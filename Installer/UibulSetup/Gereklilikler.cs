using System.Diagnostics;
using Microsoft.Win32;

namespace UibulSetup;

public enum Durum
{
    Bekliyor,
    Deneniyor,
    Tamam,       // yeşil — hazır
    Kurulacak,   // sarı — eksik ama kurulum halledebilir
    Eksik,       // kırmızı — kurulum devam edemez
    Atlandi      // gri — isteğe bağlı, kurulmadı
}

/// <summary>
/// Kurulum öncesi bakılan tek bir koşul.
/// </summary>
public sealed class Gereklilik
{
    public required string Ad { get; init; }
    public required string Aciklama { get; init; }

    /// <summary>Karşılanmazsa kurulum durur mu?</summary>
    public bool Zorunlu { get; init; } = true;

    /// <summary>Koşulu ölçer. (karşılandı mı, ayrıntı metni)</summary>
    public required Func<(bool tamam, string ayrinti)> Denetle { get; init; }

    /// <summary>Eksikse otomatik kurabiliyor muyuz? null ise kuramıyoruz.</summary>
    public Func<Action<string>, CancellationToken, Task<bool>>? Kur { get; init; }

    public Durum Sonuc { get; set; } = Durum.Bekliyor;
    public string Ayrinti { get; set; } = "";
}

/// <summary>
/// Hedef bilgisayarın UIBUL'u çalıştırıp çalıştıramayacağını ölçen kontroller.
///
/// ⚠️ TASARIM: Uygulama SELF-CONTAINED yayınlandığı için .NET runtime
/// gerekmiyor — bu yüzden .NET kontrolü "bilgi" olarak gösterilir, engel
/// değildir. Gerçekten eksik olabilecek tek şey WebView2 Runtime'dır ve
/// o da yalnızca tarayıcı algılama motorunu etkiler.
/// </summary>
public static class Gereklilikler
{
    // Microsoft'un kalıcı (evergreen) WebView2 önyükleyici bağlantısı.
    private const string WebView2Baglantisi = "https://go.microsoft.com/fwlink/p/?LinkId=2124703";

    // WebView2 Runtime'ın kayıt defterindeki ürün kimliği.
    private const string WebView2Kimlik = "{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}";

    public static List<Gereklilik> Olustur(string hedefKlasor, long gerekliBayt) => new()
    {
        new Gereklilik
        {
            Ad = "Windows sürümü",
            Aciklama = "Windows 10 veya üzeri gerekiyor",
            Denetle = () =>
            {
                var s = Environment.OSVersion.Version;
                var tamam = s.Major > 10 || (s.Major == 10 && s.Build >= 10240);
                var ad = s.Build >= 22000 ? "Windows 11" : "Windows 10";
                return (tamam, tamam ? $"{ad} (yapı {s.Build})" : $"Sürüm {s} — çok eski");
            }
        },

        new Gereklilik
        {
            Ad = "64-bit işletim sistemi",
            Aciklama = "Uygulama yalnızca 64-bit Windows'ta çalışır",
            Denetle = () => Environment.Is64BitOperatingSystem
                ? (true, "64-bit")
                : (false, "32-bit Windows desteklenmiyor")
        },

        new Gereklilik
        {
            Ad = ".NET çalışma zamanı",
            Aciklama = "Uygulamanın içine gömülü — ayrıca kurulum gerekmez",
            Denetle = () => (true, "Gerekmiyor (uygulamaya gömülü)")
        },

        new Gereklilik
        {
            Ad = "Disk alanı",
            Aciklama = "Kurulum için yeterli boş alan",
            Denetle = () =>
            {
                try
                {
                    var kok = Path.GetPathRoot(Path.GetFullPath(hedefKlasor));
                    if (string.IsNullOrEmpty(kok)) return (true, "ölçülemedi");

                    var surucu = new DriveInfo(kok);
                    var bos = surucu.AvailableFreeSpace;
                    var yeter = bos > gerekliBayt;
                    return (yeter,
                        $"{bos / 1024d / 1024 / 1024:0.0} GB boş " +
                        $"(gereken ~{gerekliBayt / 1024d / 1024:0} MB)");
                }
                catch (Exception ex) { return (true, "ölçülemedi: " + ex.Message); }
            }
        },

        new Gereklilik
        {
            Ad = "Klasör yazma izni",
            Aciklama = "Kurulum klasörüne yazılabiliyor mu",
            Denetle = () =>
            {
                try
                {
                    Directory.CreateDirectory(hedefKlasor);
                    var deneme = Path.Combine(hedefKlasor, ".yazma-denemesi");
                    File.WriteAllText(deneme, "x");
                    File.Delete(deneme);
                    return (true, "yazılabilir");
                }
                catch
                {
                    // Yönetici hakkıyla çözülebilir; kurulum bunu kendi yükseltmesiyle halleder.
                    return (false, "yazma izni yok — yönetici hakkı gerekebilir");
                }
            }
        },

        new Gereklilik
        {
            Ad = "WebView2 Runtime",
            Aciklama = "Chrome/Edge içindeki elementleri okumak için gerekli",
            Zorunlu = false,   // yoksa yalnız tarayıcı algılama motoru devre dışı kalır
            Denetle = () =>
            {
                var s = WebView2Surumu();
                return s != null ? (true, "sürüm " + s) : (false, "kurulu değil");
            },
            Kur = WebView2Kur
        },

        new Gereklilik
        {
            Ad = "Yazıcı/ekran ölçekleme",
            Aciklama = "Yüksek DPI ekranlarda görüntü kalitesi",
            Zorunlu = false,
            Denetle = () => (true, "uygulama DPI duyarlı")
        }
    };

    /// <summary>WebView2 Runtime kuruluysa sürümünü, değilse null döner.</summary>
    public static string? WebView2Surumu()
    {
        // Uc olasi konum: makine geneli (64/32 bit gorunumu) ve kullanici bazli.
        string[] yollar =
        {
            $@"SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{WebView2Kimlik}",
            $@"SOFTWARE\Microsoft\EdgeUpdate\Clients\{WebView2Kimlik}"
        };

        foreach (var yol in yollar)
        {
            var s = PvOku(Registry.LocalMachine, yol) ?? PvOku(Registry.CurrentUser, yol);
            if (!string.IsNullOrWhiteSpace(s) && s != "0.0.0.0") return s;
        }
        return null;

        static string? PvOku(RegistryKey kok, string yol)
        {
            try
            {
                using var k = kok.OpenSubKey(yol);
                return k?.GetValue("pv") as string;
            }
            catch { return null; }
        }
    }

    /// <summary>
    /// WebView2 Runtime'ı Microsoft'un resmi önyükleyicisiyle kurar.
    /// İnternet gerektirir; başarısız olursa kurulum yine de sürer.
    /// </summary>
    private static async Task<bool> WebView2Kur(Action<string> bildir, CancellationToken iptal)
    {
        var gecici = Path.Combine(Path.GetTempPath(), "MicrosoftEdgeWebview2Setup.exe");

        try
        {
            bildir("WebView2 Runtime indiriliyor (Microsoft sunucusundan)…");

            using (var http = new HttpClient { Timeout = TimeSpan.FromMinutes(5) })
            using (var akis = await http.GetStreamAsync(WebView2Baglantisi, iptal))
            using (var dosya = File.Create(gecici))
            {
                await akis.CopyToAsync(dosya, iptal);
            }

            bildir("WebView2 Runtime kuruluyor…");

            var islem = Process.Start(new ProcessStartInfo
            {
                FileName = gecici,
                Arguments = "/silent /install",
                UseShellExecute = true
            });

            if (islem == null) return false;
            await islem.WaitForExitAsync(iptal);

            var basarili = WebView2Surumu() != null;
            bildir(basarili
                ? "WebView2 Runtime kuruldu."
                : "WebView2 kurulumu tamamlanamadı (uygulama yine de çalışır).");
            return basarili;
        }
        catch (Exception ex)
        {
            bildir("WebView2 kurulamadı: " + ex.Message);
            return false;
        }
        finally
        {
            try { if (File.Exists(gecici)) File.Delete(gecici); } catch { }
        }
    }
}
