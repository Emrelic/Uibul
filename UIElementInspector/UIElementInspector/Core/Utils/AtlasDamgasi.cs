using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Tarih Atlası karesinin KİMLİĞİ — tarih, koordinat, yakınlaştırma.
    ///
    /// NEREDEN OKUNUR: tarayıcı penceresinin BAŞLIĞINDAN. Atlas sayfası
    /// document.title'ı şu hâle getirir:
    ///
    ///     Osmanlı Tarih Atlası · 1361-02-01 · 41.35N 26.50E · z6 · Edirne'nin fethi
    ///
    /// SON BÖLÜM: kronoloji sekmesinde o an açık olan MADDENİN METNİ.
    /// İsteğe bağlıdır (hiçbir madde açık değilse yoktur), ama varsa hem
    /// şeride hem panoya yazılır — çünkü "hangi maddeden bahsediyorsun"
    /// sorusu, "hangi tarih" sorusu kadar sık cevapsız kalıyor.
    ///
    /// Chrome bunun sonuna " - Google Chrome" ekler ve pencere başlığı olarak
    /// yayınlar; biz Win32'den tek çağrıyla okuruz. OCR yok, tarayıcıya
    /// bağlanma yok.
    ///
    /// ÖLÇÜLDÜ (2026-07-31): 363 karakterlik bir document.title pencere
    /// başlığında KIRPILMADAN göründü (GetWindowTextW ile 379 karakter =
    /// 363 + " - Google Chrome"). Türkçe harfler ve "·" bozulmadı. Yani
    /// bu formatın uzunluğu sorun değil.
    ///
    /// ⚠️ Başlık, kısayola basıldığı ANDA okunmalıdır — bölge seçme kaplaması
    /// açıldıktan sonra ön plandaki pencere ARTIK kaplamadır, atlas değil.
    /// </summary>
    public sealed class AtlasDamgasi
    {
        /// <summary>"1361-02-01"</summary>
        public string Tarih { get; private set; }

        /// <summary>"41.35N"</summary>
        public string Enlem { get; private set; }

        /// <summary>"26.50E"</summary>
        public string Boylam { get; private set; }

        /// <summary>"z6" — başındaki z dâhil</summary>
        public string Zoom { get; private set; }

        /// <summary>
        /// Kronoloji sekmesinde o an açık olan maddenin metni. Yoksa null.
        /// </summary>
        public string Madde { get; private set; }

        /// <summary>Damganın okunduğu ham pencere başlığı (tanı için)</summary>
        public string HamBaslik { get; private set; }

        private AtlasDamgasi() { }

        /// <summary>
        /// Görüntünün İÇİNE basılacak ve panoya metin olarak konacak satır.
        /// Dosya adı kaybolsa bile bilgi bu satırda taşınır — özelliğin
        /// varlık sebebi budur.
        /// </summary>
        public string Satir
        {
            get { return $"{Tarih} · {Enlem} {Boylam} · {Zoom} · Osmanlı Tarih Atlası"; }
        }

        /// <summary>
        /// Şeridin ikinci satırı — açık kronoloji maddesi. Madde yoksa null.
        /// </summary>
        public string MaddeSatiri
        {
            get { return string.IsNullOrWhiteSpace(Madde) ? null : "Madde: " + Madde; }
        }

        /// <summary>Panoya metin olarak konan tam kimlik (bir ya da iki satır).</summary>
        public string PanoMetni
        {
            get
            {
                var ikinci = MaddeSatiri;
                return ikinci == null ? Satir : Satir + Environment.NewLine + ikinci;
            }
        }

        /// <summary>YYYY-AA-GG_enlem_boylam_zN.png</summary>
        public string DosyaAdi
        {
            get { return $"{Tarih}_{Enlem}_{Boylam}_{Zoom}.png"; }
        }

        #region Ayrıştırma

        // Tarih: 1361-02-01 (yıl 1-5 hane; ileride MÖ için başta eksi olabilir)
        private static readonly Regex ReTarih = new Regex(
            @"(?<y>-?\d{1,5})-(?<a>\d{1,2})-(?<g>\d{1,2})", RegexOptions.Compiled);

        // Koordinat: 41.35N 26.50E  (K/G ve D/B harfleri de kabul edilir)
        private static readonly Regex ReKoordinat = new Regex(
            @"(?<la>\d{1,2}(?:[.,]\d+)?)\s*(?<lah>[NSKG])[\s,·]+(?<lo>\d{1,3}(?:[.,]\d+)?)\s*(?<loh>[EWDB])",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Yakınlaştırma: z6 | z6.2
        private static readonly Regex ReZoom = new Regex(
            @"(?<![A-Za-z0-9])z\s*(?<z>\d{1,2}(?:[.,]\d+)?)(?![0-9])",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Bir pencere başlığından damgayı çıkarır. ÜÇÜ DE bulunamazsa null
        /// döner — eksik damga basmaktansa hiç basmamak yeğdir (dosya adı
        /// kaybolduğunda kareyi kurtaracak tek şey bu satırdır).
        /// </summary>
        public static AtlasDamgasi Ayristir(string baslik)
        {
            if (string.IsNullOrWhiteSpace(baslik)) return null;

            var mT = ReTarih.Match(baslik);
            var mK = ReKoordinat.Match(baslik);
            var mZ = ReZoom.Match(baslik);
            if (!mT.Success || !mK.Success || !mZ.Success) return null;

            var yil = int.Parse(mT.Groups["y"].Value, CultureInfo.InvariantCulture);
            var ay = int.Parse(mT.Groups["a"].Value, CultureInfo.InvariantCulture);
            var gun = int.Parse(mT.Groups["g"].Value, CultureInfo.InvariantCulture);
            if (ay < 1 || ay > 12 || gun < 1 || gun > 31) return null;

            var enlem = Sayi(mK.Groups["la"].Value);
            var boylam = Sayi(mK.Groups["lo"].Value);
            var zoom = Sayi(mZ.Groups["z"].Value);
            if (double.IsNaN(enlem) || double.IsNaN(boylam) || double.IsNaN(zoom)) return null;
            if (enlem > 90 || boylam > 180) return null;

            return new AtlasDamgasi
            {
                Tarih = string.Format(CultureInfo.InvariantCulture, "{0}-{1:00}-{2:00}",
                                      yil < 0 ? yil.ToString(CultureInfo.InvariantCulture)
                                              : yil.ToString("0000", CultureInfo.InvariantCulture),
                                      ay, gun),
                Enlem = enlem.ToString("0.00", CultureInfo.InvariantCulture) + Yon(mK.Groups["lah"].Value, true),
                Boylam = boylam.ToString("0.00", CultureInfo.InvariantCulture) + Yon(mK.Groups["loh"].Value, false),
                Zoom = "z" + (zoom == Math.Floor(zoom)
                        ? ((int)zoom).ToString(CultureInfo.InvariantCulture)
                        : zoom.ToString("0.0", CultureInfo.InvariantCulture)),
                Madde = MaddeCikar(baslik),
                HamBaslik = baslik
            };
        }

        // Bir bölümün "veri" mi yoksa madde metni mi olduğunu ayırt eden desenler.
        private static readonly Regex ReSaltTarih = new Regex(@"^-?\d{1,5}-\d{1,2}-\d{1,2}$", RegexOptions.Compiled);
        private static readonly Regex ReSaltKoordinat = new Regex(
            @"^\d{1,2}(?:[.,]\d+)?\s*[NSKG][\s,]+\d{1,3}(?:[.,]\d+)?\s*[EWDB]$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        private static readonly Regex ReSaltZoom = new Regex(@"^z\s*\d{1,2}(?:[.,]\d+)?$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        // Tarayıcının başlığa eklediği kuyruk: " - Google Chrome", " — Mozilla Firefox"...
        private static readonly Regex ReTarayiciKuyrugu = new Regex(
            @"\s*[-–—]\s*(Google Chrome|Chromium|Mozilla Firefox|Microsoft.{0,3}Edge|Brave|Opera|Vivaldi)\s*$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>
        /// Başlıktaki "·" bölümlerinden, veri OLMAYAN ve uygulama adı OLMAYAN
        /// ilkini madde metni sayar. Tarayıcı kuyruğu önce atılır.
        ///
        /// ⚠️ Bu yüzden atlas tarafı madde metnindeki "·" karakterini
        /// temizlemelidir; yoksa madde iki bölüme ayrılır ve yarısı düşer.
        /// </summary>
        private static string MaddeCikar(string baslik)
        {
            var temiz = ReTarayiciKuyrugu.Replace(baslik, "");

            foreach (var ham in temiz.Split('·'))
            {
                var b = ham.Trim();
                if (b.Length == 0) continue;
                if (ReSaltTarih.IsMatch(b)) continue;
                if (ReSaltKoordinat.IsMatch(b)) continue;
                if (ReSaltZoom.IsMatch(b)) continue;
                if (b.IndexOf("atlas", StringComparison.OrdinalIgnoreCase) >= 0) continue;
                return b;
            }
            return null;
        }

        private static double Sayi(string s)
        {
            double d;
            if (double.TryParse(s.Replace(',', '.'), NumberStyles.Float,
                                CultureInfo.InvariantCulture, out d)) return d;
            return double.NaN;
        }

        // K(uzey)/G(üney) ve D(oğu)/B(atı) harflerini N/S/E/W'ye çevirir;
        // dosya adı her dilde aynı kalsın diye.
        private static string Yon(string harf, bool enlemMi)
        {
            switch (char.ToUpperInvariant(harf[0]))
            {
                case 'N': case 'K': return "N";
                case 'S': case 'G': return "S";
                case 'E': case 'D': return "E";
                case 'W': case 'B': return "W";
                default: return enlemMi ? "N" : "E";
            }
        }

        #endregion

        #region Ekrandan okuma

        /// <summary>
        /// Damgayı ekrandan okur. Sıra:
        ///   1) ÖN PLANDAKİ pencerenin başlığı (kullanıcı atlasa bakıyorken normal hâl),
        ///   2) olmazsa bütün görünür pencereler taranır ve içinde "atlas" geçen,
        ///      üç alanı da taşıyan ilk başlık alınır (Inspector ön plandayken bu çalışır).
        /// Bulunamazsa null; <paramref name="tani"/> neyin okunduğunu söyler.
        /// </summary>
        public static AtlasDamgasi Oku(out string tani)
        {
            var onPlan = BaslikAl(GetForegroundWindow());
            var d = Ayristir(onPlan);
            if (d != null)
            {
                tani = "ön plandaki pencereden okundu";
                return d;
            }

            string bulunanAtlas = null;
            foreach (var b in TumBasliklar())
            {
                if (b.IndexOf("atlas", StringComparison.OrdinalIgnoreCase) < 0) continue;
                bulunanAtlas = b;
                var d2 = Ayristir(b);
                if (d2 != null)
                {
                    tani = "arka plandaki atlas penceresinden okundu";
                    return d2;
                }
            }

            if (bulunanAtlas != null)
                tani = $"atlas penceresi bulundu ama başlıkta damga YOK: [{bulunanAtlas}]";
            else if (!string.IsNullOrWhiteSpace(onPlan))
                tani = $"atlas penceresi bulunamadı; ön plandaki başlık: [{onPlan}]";
            else
                tani = "hiçbir pencere başlığı okunamadı";
            return null;
        }

        private static IEnumerable<string> TumBasliklar()
        {
            var liste = new List<string>();
            EnumWindows((h, l) =>
            {
                if (IsWindowVisible(h))
                {
                    var t = BaslikAl(h);
                    if (!string.IsNullOrWhiteSpace(t)) liste.Add(t);
                }
                return true;
            }, IntPtr.Zero);
            return liste;
        }

        private static string BaslikAl(IntPtr hwnd)
        {
            if (hwnd == IntPtr.Zero) return null;
            int len = GetWindowTextLengthW(hwnd);
            if (len <= 0) return null;
            var sb = new StringBuilder(len + 2);
            GetWindowTextW(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private delegate bool EnumProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumProc cb, IntPtr lParam);

        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr hwnd);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextW(IntPtr hwnd, StringBuilder sb, int max);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetWindowTextLengthW(IntPtr hwnd);

        #endregion
    }
}
