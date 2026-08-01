using System.Diagnostics;
using System.IO.Compression;
using System.Reflection;
using System.Security.Principal;
using Microsoft.Win32;

namespace UibulSetup;

public sealed class KurulumSecenekleri
{
    public string HedefKlasor { get; set; } = Kurulum.VarsayilanKlasor();
    public bool MasaustuKisayolu { get; set; } = true;
    public bool BaslatMenusuKisayolu { get; set; } = true;
    public bool WindowsIleBaslat { get; set; }
    public bool KurulumSonrasiCalistir { get; set; } = true;
}

/// <summary>
/// Dosyaları yerleştiren, kısayolları ve kayıt defteri girişlerini oluşturan
/// kurulum motoru. Arayüzden bağımsızdır — sessiz kurulumda da aynısı çalışır.
/// </summary>
public static class Kurulum
{
    public const string UygulamaAdi = "UIBUL - UI Element Inspector";
    public const string KisaAd = "UIBUL";
    public const string ExeAdi = "UIElementInspector.exe";
    public const string KayitAnahtari = "UIBUL";
    public const string Yayinci = "Emrelic";
    public const string DepoAdresi = "https://github.com/Emrelic/Uibul";

    /// <summary>Gömülü yükün kaynak adı (csproj'daki LogicalName ile aynı olmalı).</summary>
    private const string YukKaynakAdi = "UIBUL_PAYLOAD";

    public static string SetupSurumu =>
        Assembly.GetExecutingAssembly().GetName().Version?.ToString(3) ?? "3.1.0";

    // ── Konumlar ──────────────────────────────────────────────────────────────

    /// <summary>
    /// Varsayılan kurulum yeri: %LOCALAPPDATA%\UIBUL.
    ///
    /// ⚠️ Program Files DEĞİL, bilinçli olarak. Oraya kurmak yönetici hakkı
    /// ister; kullanıcının kendi klasörüne kurmak UAC ekranını tamamen
    /// ortadan kaldırır. Güncellemeler de yönetici sormadan yapılabilir —
    /// otomatik güncellemenin sorunsuz çalışması buna bağlı.
    /// </summary>
    public static string VarsayilanKlasor() => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        KisaAd);

    /// <summary>Daha önce kurulmuşsa kurulu olduğu klasörü döner.</summary>
    public static string? MevcutKurulumKlasoru()
    {
        foreach (var (kok, gorunum) in new[]
                 {
                     (RegistryHive.CurrentUser, RegistryView.Default),
                     (RegistryHive.LocalMachine, RegistryView.Registry64),
                     (RegistryHive.LocalMachine, RegistryView.Registry32)
                 })
        {
            try
            {
                using var taban = RegistryKey.OpenBaseKey(kok, gorunum);
                using var k = taban.OpenSubKey(
                    $@"Software\Microsoft\Windows\CurrentVersion\Uninstall\{KayitAnahtari}");
                if (k?.GetValue("InstallLocation") is string yol &&
                    !string.IsNullOrWhiteSpace(yol) &&
                    File.Exists(Path.Combine(yol, ExeAdi)))
                {
                    return yol;
                }
            }
            catch { /* okuyamadığımız kök varsa sıradakine bak */ }
        }
        return null;
    }

    public static string? KuruluSurum()
    {
        try
        {
            using var taban = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Default);
            using var k = taban.OpenSubKey(
                $@"Software\Microsoft\Windows\CurrentVersion\Uninstall\{KayitAnahtari}");
            return k?.GetValue("DisplayVersion") as string;
        }
        catch { return null; }
    }

    // ── Yetki ─────────────────────────────────────────────────────────────────

    public static bool YoneticiMi()
    {
        try
        {
            using var kimlik = WindowsIdentity.GetCurrent();
            return new WindowsPrincipal(kimlik).IsInRole(WindowsBuiltInRole.Administrator);
        }
        catch { return false; }
    }

    /// <summary>
    /// Hedef klasör yönetici hakkı istiyor mu? Yazma denemesiyle ölçülür —
    /// yol tahmininden ("Program Files içeriyor mu") daha güvenilir.
    /// </summary>
    public static bool YoneticiGerekliMi(string hedef)
    {
        if (YoneticiMi()) return false;
        try
        {
            Directory.CreateDirectory(hedef);
            var deneme = Path.Combine(hedef, ".yetki-denemesi");
            File.WriteAllText(deneme, "x");
            File.Delete(deneme);
            return false;
        }
        catch { return true; }
    }

    /// <summary>Kurulumu yönetici olarak yeniden başlatır.</summary>
    public static bool KendiniYukselt(string argumanlar)
    {
        try
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = Environment.ProcessPath ?? Assembly.GetExecutingAssembly().Location,
                Arguments = argumanlar,
                UseShellExecute = true,
                Verb = "runas"
            });
            return true;
        }
        catch { return false; }   // kullanıcı UAC'yi reddetti
    }

    // ── Yük ───────────────────────────────────────────────────────────────────

    public static bool YukVarMi() => YukAc() != null;

    private static Stream? YukAc() =>
        Assembly.GetExecutingAssembly().GetManifestResourceStream(YukKaynakAdi);

    /// <summary>Kurulacak dosyaların açılmış toplam boyutu (yaklaşık).</summary>
    public static long YukBoyutu()
    {
        try
        {
            using var akis = YukAc();
            if (akis == null) return 0;
            using var zip = new ZipArchive(akis, ZipArchiveMode.Read);
            return zip.Entries.Sum(g => g.Length);
        }
        catch { return 0; }
    }

    // ── Kurulum ───────────────────────────────────────────────────────────────

    /// <param name="bildir">(metin, yüzde) — yüzde -1 ise belirsiz.</param>
    public static async Task Kur(
        KurulumSecenekleri secenek,
        Action<string, int> bildir,
        CancellationToken iptal = default)
    {
        bildir("Çalışan sürüm kontrol ediliyor…", 0);
        await CalisanSurumuKapat(iptal);

        bildir("Kurulum klasörü hazırlanıyor…", 3);
        Directory.CreateDirectory(secenek.HedefKlasor);

        // ── Dosyaları aç ──
        using (var akis = YukAc())
        {
            if (akis == null)
                throw new InvalidOperationException(
                    "Kurulum dosyasının içine uygulama paketi gömülmemiş. " +
                    "Bu setup hatalı üretilmiş; yeni bir kopya indirin.");

            using var zip = new ZipArchive(akis, ZipArchiveMode.Read);
            var toplam = zip.Entries.Count;
            var sayac = 0;

            foreach (var girdi in zip.Entries)
            {
                iptal.ThrowIfCancellationRequested();
                sayac++;

                // Klasör girdisi
                if (string.IsNullOrEmpty(girdi.Name))
                {
                    Directory.CreateDirectory(Path.Combine(secenek.HedefKlasor, girdi.FullName));
                    continue;
                }

                var hedef = Path.GetFullPath(Path.Combine(secenek.HedefKlasor, girdi.FullName));

                // Zip-slip koruması: hedef, kurulum klasörünün dışına çıkamaz.
                var kokTam = Path.GetFullPath(secenek.HedefKlasor);
                if (!hedef.StartsWith(kokTam, StringComparison.OrdinalIgnoreCase))
                    continue;

                var klasor = Path.GetDirectoryName(hedef);
                if (!string.IsNullOrEmpty(klasor)) Directory.CreateDirectory(klasor);

                await CopyalaYenidenDene(girdi, hedef, iptal);

                var yuzde = 5 + (int)(sayac * 80.0 / Math.Max(1, toplam));
                if (sayac % 10 == 0 || sayac == toplam)
                    bildir($"Dosyalar açılıyor… ({sayac}/{toplam})", yuzde);
            }
        }

        var exeYolu = Path.Combine(secenek.HedefKlasor, ExeAdi);
        var ikonYolu = Path.Combine(secenek.HedefKlasor, "Resources", "app.ico");
        if (!File.Exists(ikonYolu)) ikonYolu = exeYolu;

        // ── Kısayollar ──
        bildir("Kısayollar oluşturuluyor…", 88);

        if (secenek.MasaustuKisayolu)
        {
            KisayolYaz(
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                             UygulamaAdi + ".lnk"),
                exeYolu, secenek.HedefKlasor, ikonYolu);
        }

        if (secenek.BaslatMenusuKisayolu)
        {
            var menu = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Programs), KisaAd);
            Directory.CreateDirectory(menu);
            KisayolYaz(Path.Combine(menu, UygulamaAdi + ".lnk"),
                       exeYolu, secenek.HedefKlasor, ikonYolu);
        }

        // ── Windows ile başlat ──
        bildir("Başlangıç ayarı uygulanıyor…", 90);
        WindowsIleBaslatAyarla(secenek.WindowsIleBaslat, exeYolu);

        // ── Kaldırma ──
        bildir("Kaldırma bilgileri yazılıyor…", 93);
        var kaldirScript = KaldirmaScriptiYaz(secenek.HedefKlasor);
        KayitDefterineYaz(secenek.HedefKlasor, exeYolu, ikonYolu, kaldirScript);

        // ── Belgeler ──
        bildir("Belgeler yerleştiriliyor…", 97);
        BelgeKisayoluYaz(secenek.HedefKlasor);

        bildir("Kurulum tamamlandı.", 100);
    }

    /// <summary>
    /// Dosya kilitli olabilir (antivirüs tarıyor, eski süreç henüz kapanmadı).
    /// Üç kez, artan beklemeyle dener.
    /// </summary>
    private static async Task CopyalaYenidenDene(ZipArchiveEntry girdi, string hedef, CancellationToken iptal)
    {
        for (var deneme = 1; ; deneme++)
        {
            try
            {
                girdi.ExtractToFile(hedef, overwrite: true);
                return;
            }
            catch (IOException) when (deneme < 3)
            {
                await Task.Delay(400 * deneme, iptal);
            }
            catch (UnauthorizedAccessException) when (deneme < 3)
            {
                await Task.Delay(400 * deneme, iptal);
            }
        }
    }

    /// <summary>
    /// Çalışan UIBUL varsa kibarca kapatır, olmazsa sonlandırır.
    /// Güncelleme sırasında exe'nin üzerine yazılabilmesi için şart.
    /// </summary>
    public static async Task CalisanSurumuKapat(CancellationToken iptal = default)
    {
        try
        {
            var ad = Path.GetFileNameWithoutExtension(ExeAdi);
            foreach (var p in Process.GetProcessesByName(ad))
            {
                try
                {
                    if (!p.CloseMainWindow()) p.Kill();
                    if (!p.WaitForExit(5000)) p.Kill();
                    await p.WaitForExitAsync(iptal);
                }
                catch { /* zaten kapanmışsa sorun yok */ }
            }
        }
        catch { }
    }

    /// <summary>
    /// Belirtilen süreç kapanana kadar bekler (güncelleme kipinde kullanılır).
    /// </summary>
    public static async Task SureciBekle(int pid, TimeSpan zamanAsimi)
    {
        try
        {
            var p = Process.GetProcessById(pid);
            using var iptal = new CancellationTokenSource(zamanAsimi);
            await p.WaitForExitAsync(iptal.Token);
        }
        catch { /* süreç zaten yoksa ya da süre dolduysa devam et */ }
    }

    // ── Kısayol ───────────────────────────────────────────────────────────────

    /// <summary>
    /// .lnk oluşturur. Önce COM (WScript.Shell), olmazsa PowerShell.
    /// İkisi de olmazsa sessizce vazgeçer — kısayol yokluğu kurulumu bozmamalı.
    /// </summary>
    private static void KisayolYaz(string lnkYolu, string hedefExe, string calismaKlasoru, string ikon)
    {
        try
        {
            var tur = Type.GetTypeFromProgID("WScript.Shell");
            if (tur != null)
            {
                dynamic kabuk = Activator.CreateInstance(tur)!;
                dynamic kisayol = kabuk.CreateShortcut(lnkYolu);
                kisayol.TargetPath = hedefExe;
                kisayol.WorkingDirectory = calismaKlasoru;
                kisayol.Description = UygulamaAdi;
                kisayol.IconLocation = ikon + ",0";
                kisayol.Save();
                return;
            }
        }
        catch { /* COM yoksa PowerShell'e düş */ }

        try
        {
            var betik =
                $"$s=(New-Object -ComObject WScript.Shell).CreateShortcut('{lnkYolu}');" +
                $"$s.TargetPath='{hedefExe}';$s.WorkingDirectory='{calismaKlasoru}';" +
                $"$s.IconLocation='{ikon},0';$s.Save()";

            using var p = Process.Start(new ProcessStartInfo
            {
                FileName = "powershell.exe",
                Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -Command \"{betik}\"",
                UseShellExecute = false,
                CreateNoWindow = true
            });
            p?.WaitForExit(15000);
        }
        catch { }
    }

    private static void WindowsIleBaslatAyarla(bool acik, string exeYolu)
    {
        try
        {
            using var k = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Run", writable: true);
            if (k == null) return;

            if (acik) k.SetValue(KisaAd, $"\"{exeYolu}\"");
            else if (k.GetValue(KisaAd) != null) k.DeleteValue(KisaAd, throwOnMissingValue: false);
        }
        catch { }
    }

    // ── Kayıt defteri ─────────────────────────────────────────────────────────

    private static void KayitDefterineYaz(string hedef, string exe, string ikon, string kaldirScript)
    {
        try
        {
            // Kullanıcı bazlı kurulum → HKCU. Yönetici kurulumu → HKLM.
            var kok = YoneticiMi() && !hedef.StartsWith(
                          Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                          StringComparison.OrdinalIgnoreCase)
                ? Registry.LocalMachine
                : Registry.CurrentUser;

            using var k = kok.CreateSubKey(
                $@"Software\Microsoft\Windows\CurrentVersion\Uninstall\{KayitAnahtari}");
            if (k == null) return;

            long boyutKb = 0;
            try
            {
                boyutKb = new DirectoryInfo(hedef)
                    .EnumerateFiles("*", SearchOption.AllDirectories)
                    .Sum(d => d.Length) / 1024;
            }
            catch { }

            k.SetValue("DisplayName", UygulamaAdi);
            k.SetValue("DisplayVersion", SetupSurumu);
            k.SetValue("Publisher", Yayinci);
            k.SetValue("InstallLocation", hedef);
            k.SetValue("DisplayIcon", ikon);
            k.SetValue("URLInfoAbout", DepoAdresi);
            k.SetValue("InstallDate", DateTime.Now.ToString("yyyyMMdd"));
            k.SetValue("EstimatedSize", (int)Math.Max(1, boyutKb), RegistryValueKind.DWord);
            k.SetValue("NoModify", 1, RegistryValueKind.DWord);
            k.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            k.SetValue("UninstallString",
                $"powershell.exe -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File \"{kaldirScript}\"");
        }
        catch { /* kayıt yazılamazsa program yine çalışır, sadece "Uygulamalar" listesinde görünmez */ }
    }

    /// <summary>
    /// Kaldırma betiğini yazar.
    ///
    /// ⚠️ Neden PowerShell betiği, ayrı bir kaldırma .exe'si değil:
    /// Uygulama self-contained; ayrı bir kaldırıcı exe de kendi .NET
    /// kopyasını taşımak zorunda kalır ve kuruluma 70 MB daha eklerdi.
    /// Betik, kendisini de siler.
    /// </summary>
    private static string KaldirmaScriptiYaz(string hedef)
    {
        var yol = Path.Combine(hedef, "Kaldir.ps1");

        var masaustu = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
            UygulamaAdi + ".lnk");
        var menuKlasor = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Programs), KisaAd);

        var betik = $@"# UIBUL kaldırma betiği — kurulum tarafından üretildi
Add-Type -AssemblyName System.Windows.Forms | Out-Null

$cevap = [System.Windows.Forms.MessageBox]::Show(
    ""{UygulamaAdi} kaldırılsın mı?`n`nKurulum klasörü ve kısayollar silinecek.`nAyarlarınız ve arşiviniz SİLİNMEZ."",
    ""UIBUL kaldırma"",
    [System.Windows.Forms.MessageBoxButtons]::YesNo,
    [System.Windows.Forms.MessageBoxIcon]::Question)

if ($cevap -ne [System.Windows.Forms.DialogResult]::Yes) {{ exit }}

# Çalışan kopyayı kapat
Get-Process -Name '{Path.GetFileNameWithoutExtension(ExeAdi)}' -ErrorAction SilentlyContinue |
    ForEach-Object {{ try {{ $_.Kill(); $_.WaitForExit(5000) }} catch {{ }} }}
Start-Sleep -Milliseconds 700

# Kısayollar
Remove-Item -LiteralPath '{masaustu}' -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath '{menuKlasor}' -Recurse -Force -ErrorAction SilentlyContinue

# Başlangıç kaydı
Remove-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run' `
    -Name '{KisaAd}' -Force -ErrorAction SilentlyContinue

# Kayıt defteri
Remove-Item -Path 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{KayitAnahtari}' `
    -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item -Path 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{KayitAnahtari}' `
    -Recurse -Force -ErrorAction SilentlyContinue

# Kurulum klasörü — betik kendi içinde olduğu için gecikmeli silinir
$hedef = '{hedef}'
$temizle = ""Start-Sleep -Seconds 2; Remove-Item -LiteralPath '$hedef' -Recurse -Force -ErrorAction SilentlyContinue""
Start-Process powershell -ArgumentList '-NoProfile','-WindowStyle','Hidden','-Command',$temizle

[System.Windows.Forms.MessageBox]::Show(
    ""{UygulamaAdi} kaldırıldı.`n`nAyar ve arşiv dosyalarınız duruyor:`n%AppData%\UIElementInspector"",
    ""Kaldırma tamamlandı"",
    [System.Windows.Forms.MessageBoxButtons]::OK,
    [System.Windows.Forms.MessageBoxIcon]::Information) | Out-Null
";

        File.WriteAllText(yol, betik, new System.Text.UTF8Encoding(true));
        return yol;
    }

    /// <summary>Başlat menüsüne belgeler için de bir kısayol koyar.</summary>
    private static void BelgeKisayoluYaz(string hedef)
    {
        try
        {
            var belge = Path.Combine(hedef, "Docs", "TANITIM.html");
            if (!File.Exists(belge)) return;

            var menu = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Programs), KisaAd);
            Directory.CreateDirectory(menu);

            File.WriteAllText(Path.Combine(menu, "UIBUL Tanıtım ve Kılavuz.url"),
                $"[InternetShortcut]{Environment.NewLine}URL=file:///{belge.Replace('\\', '/')}{Environment.NewLine}");
        }
        catch { }
    }
}
