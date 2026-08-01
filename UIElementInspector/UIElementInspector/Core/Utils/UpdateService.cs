using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Newtonsoft.Json.Linq;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Tek bir yayınlanmış sürümün bilgisi (GitHub release karşılığı).
    /// </summary>
    public sealed class SurumBilgisi
    {
        public Version Surum { get; set; } = new Version(0, 0, 0);
        public string Etiket { get; set; } = "";
        public string Baslik { get; set; } = "";
        public string Notlar { get; set; } = "";
        public string IndirmeAdresi { get; set; } = "";
        public string DosyaAdi { get; set; } = "";
        public long Boyut { get; set; }
        public DateTimeOffset Tarih { get; set; }

        public string BoyutMetni => Boyut <= 0
            ? "bilinmiyor"
            : $"{Boyut / 1024d / 1024d:0.#} MB";
    }

    /// <summary>
    /// Güncelleme kontrolünün sonucu. Kontrol başarısız olduysa
    /// <see cref="Hata"/> doludur ve <see cref="GuncellemeVar"/> false döner —
    /// ağ hatası asla "güncelleme yok" gibi gösterilmez.
    /// </summary>
    public sealed class GuncellemeSonucu
    {
        public bool Basarili { get; set; }
        public string? Hata { get; set; }
        public Version MevcutSurum { get; set; } = new Version(0, 0, 0);
        public SurumBilgisi? Yeni { get; set; }

        public bool GuncellemeVar =>
            Basarili && Yeni != null && Yeni.Surum > MevcutSurum;

        /// <summary>Yeni sürüm var ama indirilecek setup dosyası eklenmemiş.</summary>
        public bool DosyaEksik =>
            GuncellemeVar && string.IsNullOrWhiteSpace(Yeni!.IndirmeAdresi);
    }

    /// <summary>
    /// GitHub Releases üzerinden sürüm kontrolü ve güncelleme indirme.
    ///
    /// ⚠️ TASARIM NOTU — neden GitHub API'si:
    /// Depo (github.com/Emrelic/Uibul) zaten var ve public. Releases API'si
    /// kimlik doğrulama istemez, saatte 60 istek hakkı verir; uygulama günde
    /// bir kez baktığı için bu limit hiç zorlanmaz. Arkadaşınızın token
    /// girmesine, hesap açmasına gerek yoktur.
    ///
    /// ⚠️ Sürüm karşılaştırması ETİKETTEN yapılır ("v3.2.0" → 3.2.0), release
    /// başlığından değil. Başlık serbest metindir, etiket ise sürüm numarası
    /// olmak zorundadır.
    /// </summary>
    public static class UpdateService
    {
        /// <summary>Varsayılan depo. Ayarlardan değiştirilebilir.</summary>
        public const string VarsayilanDepo = "Emrelic/Uibul";

        private const string KullaniciAjani = "UIBUL-UpdateChecker";

        private static readonly HttpClient _http = OlusturHttpClient();

        private static HttpClient OlusturHttpClient()
        {
            var c = new HttpClient { Timeout = TimeSpan.FromMinutes(10) };
            // GitHub API User-Agent olmadan 403 döner — bu zorunlu.
            c.DefaultRequestHeaders.UserAgent.ParseAdd(KullaniciAjani);
            c.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/vnd.github+json"));
            return c;
        }

        /// <summary>
        /// Çalışan uygulamanın sürümü. csproj'daki &lt;Version&gt; alanından gelir.
        /// </summary>
        public static Version MevcutSurum
        {
            get
            {
                try
                {
                    var asm = Assembly.GetExecutingAssembly();

                    // InformationalVersion "3.1.0+abc123" gibi olabilir; + sonrası atılır.
                    var bilgi = asm.GetCustomAttribute<AssemblyInformationalVersionAttribute>()
                                   ?.InformationalVersion;
                    if (!string.IsNullOrWhiteSpace(bilgi))
                    {
                        var temiz = bilgi.Split('+')[0].Split('-')[0];
                        if (Version.TryParse(temiz, out var v1)) return Normalle(v1);
                    }

                    var v2 = asm.GetName().Version;
                    if (v2 != null) return Normalle(v2);
                }
                catch { /* sürüm okunamazsa 0.0.0 döner, her release yeni sayılır */ }

                return new Version(0, 0, 0);
            }
        }

        public static string MevcutSurumMetni => MevcutSurum.ToString(3);

        /// <summary>
        /// Version'ı 3 haneye indirir. AssemblyVersion 4 haneli (3.1.0.0),
        /// etiketler 3 haneli (3.1.0) olduğu için karşılaştırma bozulmasın diye.
        /// </summary>
        private static Version Normalle(Version v) =>
            new Version(Math.Max(v.Major, 0), Math.Max(v.Minor, 0), Math.Max(v.Build, 0));

        /// <summary>
        /// "v3.2.0", "3.2", "sürüm-3.2.0" gibi etiketlerden sürüm çıkarır.
        /// Çıkaramazsa null döner.
        /// </summary>
        public static Version? EtiketiCozumle(string? etiket)
        {
            if (string.IsNullOrWhiteSpace(etiket)) return null;

            var rakamlar = new string(etiket.Select(k => char.IsDigit(k) || k == '.' ? k : ' ').ToArray());
            foreach (var parca in rakamlar.Split(' ', StringSplitOptions.RemoveEmptyEntries))
            {
                var t = parca.Trim('.');
                if (t.Length == 0) continue;

                // "3" tek başına Version.TryParse'a takılır; ".0" ekleyerek düzelt.
                var aday = t.Contains('.') ? t : t + ".0";
                if (Version.TryParse(aday, out var v)) return Normalle(v);
            }
            return null;
        }

        /// <summary>
        /// En son yayınlanan sürümü sorar ve mevcutla karşılaştırır.
        /// Ağ hatası fırlatmaz; sonucu <see cref="GuncellemeSonucu.Hata"/>'ya yazar.
        /// </summary>
        public static async Task<GuncellemeSonucu> KontrolEtAsync(
            string? depo = null, CancellationToken iptal = default)
        {
            var sonuc = new GuncellemeSonucu { MevcutSurum = MevcutSurum };
            depo = string.IsNullOrWhiteSpace(depo) ? VarsayilanDepo : depo!.Trim();

            try
            {
                var adres = $"https://api.github.com/repos/{depo}/releases/latest";
                using var cevap = await _http.GetAsync(adres, iptal).ConfigureAwait(false);

                if (cevap.StatusCode == System.Net.HttpStatusCode.NotFound)
                {
                    sonuc.Hata = $"'{depo}' deposunda henüz yayınlanmış bir sürüm yok.";
                    return sonuc;
                }

                if (!cevap.IsSuccessStatusCode)
                {
                    sonuc.Hata = $"GitHub yanıtı: {(int)cevap.StatusCode} {cevap.ReasonPhrase}";
                    return sonuc;
                }

                var govde = await cevap.Content.ReadAsStringAsync(iptal).ConfigureAwait(false);
                var j = JObject.Parse(govde);

                var etiket = j["tag_name"]?.ToString() ?? "";
                var surum = EtiketiCozumle(etiket);
                if (surum == null)
                {
                    sonuc.Hata = $"Sürüm etiketi anlaşılamadı: '{etiket}'";
                    return sonuc;
                }

                var bilgi = new SurumBilgisi
                {
                    Surum = surum,
                    Etiket = etiket,
                    Baslik = j["name"]?.ToString() ?? etiket,
                    Notlar = j["body"]?.ToString() ?? "",
                    Tarih = j["published_at"]?.ToObject<DateTimeOffset>() ?? DateTimeOffset.MinValue
                };

                // Kurulum dosyasını seç: .exe tercih, yoksa .zip.
                var varliklar = j["assets"] as JArray ?? new JArray();
                var secilen =
                    varliklar.FirstOrDefault(a => (a["name"]?.ToString() ?? "")
                        .EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                    ?? varliklar.FirstOrDefault(a => (a["name"]?.ToString() ?? "")
                        .EndsWith(".zip", StringComparison.OrdinalIgnoreCase));

                if (secilen != null)
                {
                    bilgi.DosyaAdi = secilen["name"]?.ToString() ?? "UIBUL_Setup.exe";
                    bilgi.IndirmeAdresi = secilen["browser_download_url"]?.ToString() ?? "";
                    bilgi.Boyut = secilen["size"]?.ToObject<long>() ?? 0;
                }

                sonuc.Yeni = bilgi;
                sonuc.Basarili = true;
                return sonuc;
            }
            catch (OperationCanceledException)
            {
                sonuc.Hata = "Kontrol iptal edildi.";
                return sonuc;
            }
            catch (HttpRequestException ex)
            {
                sonuc.Hata = $"İnternete ulaşılamadı: {ex.Message}";
                return sonuc;
            }
            catch (Exception ex)
            {
                sonuc.Hata = ex.Message;
                return sonuc;
            }
        }

        /// <summary>
        /// Kurulum dosyasını indirir, tam yolunu döner.
        /// İlerleme: (indirilen bayt, toplam bayt) — toplam bilinmiyorsa -1.
        /// </summary>
        public static async Task<string> IndirAsync(
            SurumBilgisi bilgi,
            IProgress<(long inen, long toplam)>? ilerleme = null,
            CancellationToken iptal = default)
        {
            if (bilgi == null) throw new ArgumentNullException(nameof(bilgi));
            if (string.IsNullOrWhiteSpace(bilgi.IndirmeAdresi))
                throw new InvalidOperationException("Bu sürüme kurulum dosyası eklenmemiş.");

            var klasor = Path.Combine(Path.GetTempPath(), "UIBUL_Guncelleme");
            Directory.CreateDirectory(klasor);

            var dosyaAdi = string.IsNullOrWhiteSpace(bilgi.DosyaAdi) ? "UIBUL_Setup.exe" : bilgi.DosyaAdi;
            var hedef = Path.Combine(klasor, dosyaAdi);

            // Yarım kalmış indirmeyi temizle.
            if (File.Exists(hedef)) File.Delete(hedef);
            var gecici = hedef + ".indiriliyor";
            if (File.Exists(gecici)) File.Delete(gecici);

            using (var cevap = await _http.GetAsync(
                       bilgi.IndirmeAdresi, HttpCompletionOption.ResponseHeadersRead, iptal)
                   .ConfigureAwait(false))
            {
                cevap.EnsureSuccessStatusCode();

                var toplam = cevap.Content.Headers.ContentLength ?? bilgi.Boyut;
                if (toplam <= 0) toplam = -1;

                using var kaynak = await cevap.Content.ReadAsStreamAsync(iptal).ConfigureAwait(false);
                using var yazici = new FileStream(gecici, FileMode.Create, FileAccess.Write,
                                                  FileShare.None, 81920, useAsync: true);

                var tampon = new byte[81920];
                long inen = 0;
                int okunan;
                while ((okunan = await kaynak.ReadAsync(tampon, iptal).ConfigureAwait(false)) > 0)
                {
                    await yazici.WriteAsync(tampon.AsMemory(0, okunan), iptal).ConfigureAwait(false);
                    inen += okunan;
                    ilerleme?.Report((inen, toplam));
                }
            }

            File.Move(gecici, hedef);
            return hedef;
        }

        /// <summary>
        /// İndirilen setup'ı güncelleme kipinde başlatır ve uygulamadan çıkar.
        ///
        /// ⚠️ Uygulama KAPANMAK ZORUNDA: kendi .exe'sinin üzerine yazılacak.
        /// Setup, /GUNCELLEME bayrağıyla eski sürümün kapanmasını bekler,
        /// dosyaları değiştirir ve programı yeniden başlatır.
        /// </summary>
        public static void KurVeCik(string setupYolu)
        {
            if (!File.Exists(setupYolu))
                throw new FileNotFoundException("Kurulum dosyası bulunamadı.", setupYolu);

            var kendiPid = Process.GetCurrentProcess().Id;

            var baslat = new ProcessStartInfo
            {
                FileName = setupYolu,
                Arguments = $"/GUNCELLEME /BEKLE:{kendiPid}",
                UseShellExecute = true
            };
            Process.Start(baslat);

            System.Windows.Application.Current?.Shutdown();
        }

        /// <summary>Depo sayfasını tarayıcıda açar (elle indirme için).</summary>
        public static void DepoyuAc(string? depo = null)
        {
            depo = string.IsNullOrWhiteSpace(depo) ? VarsayilanDepo : depo!.Trim();
            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = $"https://github.com/{depo}/releases/latest",
                    UseShellExecute = true
                });
            }
            catch { /* tarayıcı yoksa sessiz geç */ }
        }
    }
}
