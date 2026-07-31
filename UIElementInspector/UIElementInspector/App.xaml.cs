using System;
using System.Configuration;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows;

namespace UIElementInspector
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : System.Windows.Application
    {
        private static Mutex _tekOrnek;

        /// <summary>
        /// ⚠️ TEK ÖRNEK ZORUNLU — kısayollar yüzünden.
        ///
        /// Global kısayolu (RegisterHotKey) yalnız İLK örnek alabilir. İkinci
        /// örnekte kayıt 1409 ile başarısız olur ve HotkeyService düşük
        /// seviyeli klavye kancasına düşer; o kanca tuşu YUTMADIĞI için aynı
        /// tuşa iki örnek birden tepki verir, iki kaplama üst üste açılır ve
        /// hangisinin kare aldığı belirsizleşir.
        ///
        /// Yaşandı: masaüstü kısayolundan ikinci kez açılınca iki örnek
        /// (pid 18360 ve 23428) aynı anda çalışıyordu.
        /// </summary>
        /// <summary>
        /// Çalışan örneğe "kendini göster" demek için yayınlanan ileti.
        ///
        /// ⚠️ NEDEN PENCERE TUTAMACI YETMİYOR: uygulama incelemeye başlarken
        /// ve kare alırken ana pencereyi GİZLİYOR. Gizli bir WPF penceresi için
        /// Process.MainWindowHandle SIFIR döner. İlk sürüm bu yüzden gizli
        /// örneği "penceresi yok" sayıp atlıyordu; ikinci kopya da tek örnek
        /// kuralı gereği kapandığı için kısayola basınca HİÇBİR ŞEY OLMUYORDU.
        /// İleti yayını tutamaçtan bağımsızdır, gizli pencereye de ulaşır.
        /// </summary>
        public static readonly uint GOSTER_ILETISI =
            RegisterWindowMessage("UIElementInspector_PencereyiGoster");

        protected override void OnStartup(StartupEventArgs e)
        {
            bool ilkMi;
            _tekOrnek = new Mutex(true, @"Global\UIElementInspector_TekOrnek", out ilkMi);

            if (!ilkMi)
            {
                // Önce yayın: gizli pencereli örnek de duyar.
                PostMessage(HWND_BROADCAST, GOSTER_ILETISI, IntPtr.Zero, IntPtr.Zero);
                // Sonra klasik yol: görünür pencere varsa öne gelsin.
                OncekiniOneGetir();
                Shutdown();
                return;
            }

            base.OnStartup(e);
        }

        protected override void OnExit(ExitEventArgs e)
        {
            try { _tekOrnek?.ReleaseMutex(); } catch { }
            _tekOrnek?.Dispose();
            base.OnExit(e);
        }

        private static void OncekiniOneGetir()
        {
            try
            {
                var benim = Process.GetCurrentProcess();
                foreach (var p in Process.GetProcessesByName(benim.ProcessName))
                {
                    if (p.Id == benim.Id) continue;
                    if (p.MainWindowHandle == IntPtr.Zero) continue;

                    ShowWindow(p.MainWindowHandle, SW_RESTORE);
                    SetForegroundWindow(p.MainWindowHandle);
                    break;
                }
            }
            catch { /* öne getiremezsek de yeni örnek açılmamalı */ }
        }

        private const int SW_RESTORE = 9;
        private static readonly IntPtr HWND_BROADCAST = new IntPtr(0xFFFF);

        [DllImport("user32.dll")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        private static extern uint RegisterWindowMessage(string lpString);

        [DllImport("user32.dll")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);
    }
}
