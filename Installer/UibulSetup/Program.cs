namespace UibulSetup;

internal static class Program
{
    /// <summary>
    /// Kurulum programının giriş noktası.
    ///
    /// Desteklenen komut satırı seçenekleri:
    ///   (yok)              Normal sihirbaz
    ///   /SESSIZ            Soru sormadan varsayılan ayarlarla kur
    ///   /GUNCELLEME        Güncelleme kipi — mevcut kuruluma sessizce yaz, sonra çalıştır
    ///   /BEKLE:1234        Bu süreç kapanana kadar bekle (güncellemede uygulamanın kendisi)
    ///   /KLASOR:C:\yol     Kurulum klasörünü belirt
    ///   /YUKSELTILDI       İç kullanım: yönetici olarak yeniden başlatıldığını belirtir
    /// </summary>
    [STAThread]
    private static void Main(string[] argumanlar)
    {
        ApplicationConfiguration.Initialize();

        var sessiz = Bayrak(argumanlar, "/SESSIZ");
        var guncelleme = Bayrak(argumanlar, "/GUNCELLEME");
        var beklenecekPid = SayiDeger(argumanlar, "/BEKLE:");
        var klasor = MetinDeger(argumanlar, "/KLASOR:");

        // Güncelleme kipi: uygulamanın kapanmasını bekle, mevcut yere kur, geri aç.
        if (guncelleme)
        {
            GuncellemeKipi(beklenecekPid, klasor).GetAwaiter().GetResult();
            return;
        }

        if (sessiz)
        {
            SessizKurulum(klasor).GetAwaiter().GetResult();
            return;
        }

        Application.Run(new SetupForm(klasor));
    }

    /// <summary>
    /// Uygulamanın içinden tetiklenen güncelleme. Arayüz göstermez; yalnızca
    /// hata olursa bir kutu çıkarır — sessizce başarısız olup kullanıcıyı
    /// eski sürümde bırakmak en kötü sonuç olurdu.
    /// </summary>
    private static async Task GuncellemeKipi(int beklenecekPid, string? klasor)
    {
        try
        {
            if (beklenecekPid > 0)
                await Kurulum.SureciBekle(beklenecekPid, TimeSpan.FromSeconds(30));

            await Kurulum.CalisanSurumuKapat();

            var hedef = klasor
                        ?? Kurulum.MevcutKurulumKlasoru()
                        ?? Kurulum.VarsayilanKlasor();

            // Güncellemede kısayolları yeniden yazmak zararsızdır (yol değişmiş
            // olabilir), ama "Windows ile başlat" tercihine dokunulmaz.
            var secenek = new KurulumSecenekleri
            {
                HedefKlasor = hedef,
                MasaustuKisayolu = true,
                BaslatMenusuKisayolu = true,
                WindowsIleBaslat = WindowsIleBaslatAcikMi(),
                KurulumSonrasiCalistir = true
            };

            await Kurulum.Kur(secenek, (_, _) => { });

            var exe = Path.Combine(hedef, Kurulum.ExeAdi);
            if (File.Exists(exe))
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = exe,
                    WorkingDirectory = hedef,
                    UseShellExecute = true
                });
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                "Güncelleme tamamlanamadı.\n\n" + ex.Message +
                "\n\nProgramın eski sürümü çalışmaya devam eder. " +
                "Kurulum dosyasını elle çalıştırarak güncelleyebilirsiniz.",
                "UIBUL güncelleme", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }

    private static async Task SessizKurulum(string? klasor)
    {
        try
        {
            var secenek = new KurulumSecenekleri
            {
                HedefKlasor = klasor ?? Kurulum.VarsayilanKlasor()
            };
            await Kurulum.Kur(secenek, (_, _) => { });
        }
        catch (Exception ex)
        {
            MessageBox.Show("Sessiz kurulum başarısız: " + ex.Message,
                "UIBUL kurulum", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static bool WindowsIleBaslatAcikMi()
    {
        try
        {
            using var k = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows\CurrentVersion\Run");
            return k?.GetValue(Kurulum.KisaAd) != null;
        }
        catch { return false; }
    }

    // ── Argüman yardımcıları ──────────────────────────────────────────────────

    private static bool Bayrak(string[] a, string ad) =>
        a.Any(x => string.Equals(x, ad, StringComparison.OrdinalIgnoreCase));

    private static string? MetinDeger(string[] a, string onek)
    {
        var g = a.FirstOrDefault(x => x.StartsWith(onek, StringComparison.OrdinalIgnoreCase));
        var d = g?[onek.Length..].Trim('"');
        return string.IsNullOrWhiteSpace(d) ? null : d;
    }

    private static int SayiDeger(string[] a, string onek) =>
        int.TryParse(MetinDeger(a, onek), out var s) ? s : 0;
}
