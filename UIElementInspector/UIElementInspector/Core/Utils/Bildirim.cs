using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
// WinForms de referansli oldugu icin Color/Brushes iki ad alaninda birden var.
using Color = System.Windows.Media.Color;
using Brushes = System.Windows.Media.Brushes;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Ekranın sağ altında birkaç saniye görünen bildirim.
    ///
    /// ⚠️ NEDEN VAR: F11 karesi ana pencere gizliyken alınıyor ve sonuç
    /// yalnızca uygulamanın konsol sekmesine yazılıyordu. Kullanıcı F11'e
    /// bastı, hiçbir şey görmedi, "çalışmıyor" diye bildirdi — oysa kısayol
    /// çalışıyordu, kare almayı kod reddetmişti ve red görünmüyordu.
    /// SESSİZ BAŞARISIZLIK, HATANIN KENDİSİNDEN KÖTÜDÜR.
    /// </summary>
    public static class Bildirim
    {
        public static void Goster(string baslik, string alt, bool hata = false,
                                  int sureMs = 3000)
        {
            var uygulama = System.Windows.Application.Current;
            if (uygulama == null) return;

            uygulama.Dispatcher.BeginInvoke(new Action(() =>
            {
                try { Ciz(baslik, alt, hata, sureMs); }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("Bildirim gosterilemedi: " + ex.Message);
                }
            }), DispatcherPriority.Normal);
        }

        private static void Ciz(string baslik, string alt, bool hata, int sureMs)
        {
            var zemin = hata
                ? Color.FromRgb(120, 22, 28)
                : Color.FromRgb(26, 32, 40);
            var kenar = hata
                ? Color.FromRgb(220, 60, 60)
                : Color.FromRgb(70, 130, 180);

            var yigin = new StackPanel();
            yigin.Children.Add(new TextBlock
            {
                Text = baslik,
                Foreground = Brushes.White,
                FontSize = 14,
                FontWeight = FontWeights.Bold,
                TextWrapping = TextWrapping.Wrap
            });

            if (!string.IsNullOrWhiteSpace(alt))
            {
                yigin.Children.Add(new TextBlock
                {
                    Text = alt,
                    Foreground = new SolidColorBrush(Color.FromRgb(210, 215, 225)),
                    FontSize = 12,
                    Margin = new Thickness(0, 4, 0, 0),
                    TextWrapping = TextWrapping.Wrap
                });
            }

            var pencere = new Window
            {
                WindowStyle = WindowStyle.None,
                AllowsTransparency = true,
                Background = Brushes.Transparent,
                Topmost = true,
                ShowInTaskbar = false,
                ShowActivated = false,          // odağı ÇALMAZ; kullanıcı yazmaya devam edebilir
                SizeToContent = SizeToContent.Height,
                Width = 420,
                ResizeMode = ResizeMode.NoResize,
                Content = new Border
                {
                    Background = new SolidColorBrush(zemin),
                    BorderBrush = new SolidColorBrush(kenar),
                    BorderThickness = new Thickness(2),
                    CornerRadius = new CornerRadius(6),
                    Padding = new Thickness(14, 11, 14, 11),
                    Child = yigin
                }
            };

            var alan = SystemParameters.WorkArea;
            pencere.Left = alan.Right - pencere.Width - 20;
            pencere.Show();
            pencere.Top = alan.Bottom - pencere.ActualHeight - 20;

            // Tıklayınca hemen kapansın
            pencere.MouseLeftButtonDown += (s, e) => { try { pencere.Close(); } catch { } };

            var sayac = new DispatcherTimer { Interval = TimeSpan.FromMilliseconds(sureMs) };
            sayac.Tick += (s, e) =>
            {
                sayac.Stop();
                try { pencere.Close(); } catch { }
            };
            sayac.Start();
        }
    }
}
