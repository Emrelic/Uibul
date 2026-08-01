using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

// WinForms de projede acik oldugu icin bu adlar belirsiz kaliyor; WPF'i secelim.
using Button = System.Windows.Controls.Button;
using Color = System.Windows.Media.Color;
using Brushes = System.Windows.Media.Brushes;
using HorizontalAlignment = System.Windows.HorizontalAlignment;

namespace UIElementInspector.Windows
{
    /// <summary>
    /// Adım adım öğretici. İçerik <see cref="TutorialContent"/> içindedir;
    /// bu sınıf yalnızca gezinme ve çizim yapar.
    /// </summary>
    public partial class TutorialWindow : Window
    {
        /// <summary>
        /// İlk çalıştırma işareti. Kurulum klasörüne değil kullanıcı profiline
        /// yazılır — program Program Files altındaysa oraya yazma izni olmaz.
        /// </summary>
        private static string IsaretDosyasi => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "UIElementInspector", "ogretici-gosterildi.txt");

        private readonly IReadOnlyList<Adim> _adimlar = TutorialContent.Adimlar;
        private readonly List<Button> _navDugmeleri = new();
        private int _sira;

        public TutorialWindow()
        {
            InitializeComponent();

            txtSurumEtiketi.Text = $"UIBUL v{Core.Utils.UpdateService.MevcutSurumMetni} · {_adimlar.Count} adım";
            chkTekrarGosterme.IsChecked = File.Exists(IsaretDosyasi);

            NavigasyonuKur();
            Goster(0);
        }

        // ── İlk çalıştırma ────────────────────────────────────────────────────

        public static bool IlkCalistirmaMi()
        {
            try { return !File.Exists(IsaretDosyasi); }
            catch { return false; }
        }

        public static void IlkCalistirmayiIsaretle()
        {
            try
            {
                var klasor = Path.GetDirectoryName(IsaretDosyasi);
                if (!string.IsNullOrEmpty(klasor)) Directory.CreateDirectory(klasor);
                File.WriteAllText(IsaretDosyasi,
                    "Bu dosya, öğreticinin açılışta bir kez gösterildiğini belirtir.\n" +
                    "Silerseniz öğretici bir sonraki açılışta yeniden çıkar.\n" +
                    $"Tarih: {DateTime.Now:yyyy-MM-dd HH:mm}\n");
            }
            catch { /* yazamazsak öğretici her açılışta çıkar; kırıcı değil */ }
        }

        private void TekrarGosterme_Click(object sender, RoutedEventArgs e)
        {
            if (chkTekrarGosterme.IsChecked == true) IlkCalistirmayiIsaretle();
            else { try { if (File.Exists(IsaretDosyasi)) File.Delete(IsaretDosyasi); } catch { } }
        }

        // ── Sol gezinme ───────────────────────────────────────────────────────

        private void NavigasyonuKur()
        {
            var panel = new StackPanel();
            string sonBolum = "";

            for (int i = 0; i < _adimlar.Count; i++)
            {
                var adim = _adimlar[i];

                if (adim.Bolum != sonBolum)
                {
                    sonBolum = adim.Bolum;
                    panel.Children.Add(new TextBlock
                    {
                        Text = adim.Bolum,
                        Foreground = new SolidColorBrush(Color.FromRgb(0x4F, 0xC3, 0xF7)),
                        FontSize = 10,
                        FontWeight = FontWeights.Bold,
                        Margin = new Thickness(18, 16, 18, 6)
                    });
                }

                var dugme = new Button
                {
                    Content = new TextBlock
                    {
                        Text = adim.Baslik,
                        TextWrapping = TextWrapping.Wrap,
                        FontSize = 12
                    },
                    Tag = i,
                    HorizontalContentAlignment = HorizontalAlignment.Left,
                    Padding = new Thickness(18, 7, 14, 7),
                    Margin = new Thickness(0),
                    BorderThickness = new Thickness(3, 0, 0, 0),
                    BorderBrush = Brushes.Transparent,
                    Background = Brushes.Transparent,
                    Foreground = new SolidColorBrush(Color.FromRgb(0xCF, 0xD8, 0xDC)),
                    Cursor = System.Windows.Input.Cursors.Hand
                };
                dugme.Click += (s, _) => Goster((int)((Button)s).Tag!);

                _navDugmeleri.Add(dugme);
                panel.Children.Add(dugme);
            }

            listBolumler.Items.Clear();
            listBolumler.Items.Add(panel);
        }

        private void NavVurgula()
        {
            for (int i = 0; i < _navDugmeleri.Count; i++)
            {
                var aktif = i == _sira;
                _navDugmeleri[i].Background = aktif
                    ? new SolidColorBrush(Color.FromRgb(0x37, 0x47, 0x4F))
                    : Brushes.Transparent;
                _navDugmeleri[i].BorderBrush = aktif
                    ? new SolidColorBrush(Color.FromRgb(0x4C, 0xAF, 0x50))
                    : Brushes.Transparent;
                _navDugmeleri[i].Foreground = aktif
                    ? Brushes.White
                    : new SolidColorBrush(Color.FromRgb(0xCF, 0xD8, 0xDC));
            }
        }

        // ── Adım gösterimi ────────────────────────────────────────────────────

        private void Goster(int sira)
        {
            if (sira < 0 || sira >= _adimlar.Count) return;
            _sira = sira;
            var adim = _adimlar[sira];

            txtBolumEtiketi.Text = adim.Bolum;
            txtBaslik.Text = adim.Baslik;
            txtOzet.Text = adim.Ozet;

            pnlIcerik.Children.Clear();
            foreach (var blok in adim.Bloklar)
                pnlIcerik.Children.Add(BlokCiz(blok));

            scrollIcerik.ScrollToTop();

            btnGeri.IsEnabled = sira > 0;
            btnIleri.Content = sira == _adimlar.Count - 1 ? "Bitir ✓" : "İleri ▶";

            var yuzde = (sira + 1) * 100.0 / _adimlar.Count;
            pbIlerleme.Value = yuzde;
            txtIlerlemeMetni.Text = $"Adım {sira + 1} / {_adimlar.Count}";

            NavVurgula();
        }

        private UIElement BlokCiz(Blok blok) => blok.Tur switch
        {
            BlokTuru.Baslik => new TextBlock
            {
                Text = blok.Metin,
                FontSize = 15,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(Color.FromRgb(0x15, 0x65, 0xC0)),
                Margin = new Thickness(0, 18, 0, 8),
                TextWrapping = TextWrapping.Wrap
            },

            BlokTuru.Paragraf => new TextBlock
            {
                Text = blok.Metin,
                FontSize = 13.5,
                LineHeight = 22,
                Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33)),
                Margin = new Thickness(0, 0, 0, 12),
                TextWrapping = TextWrapping.Wrap
            },

            BlokTuru.Madde => ListeCiz(blok.Ogeler, numarali: false),
            BlokTuru.Numarali => ListeCiz(blok.Ogeler, numarali: true),
            BlokTuru.Tus => TusCiz(blok.Etiket, blok.Metin),

            BlokTuru.Ipucu => KutuCiz("💡 İPUCU", blok.Metin,
                Color.FromRgb(0xE8, 0xF5, 0xE9), Color.FromRgb(0x2E, 0x7D, 0x32)),
            BlokTuru.Uyari => KutuCiz("⚠️ DİKKAT", blok.Metin,
                Color.FromRgb(0xFF, 0xF3, 0xE0), Color.FromRgb(0xE6, 0x51, 0x00)),
            BlokTuru.Bilgi => KutuCiz("ℹ️ BİLGİ", blok.Metin,
                Color.FromRgb(0xE3, 0xF2, 0xFD), Color.FromRgb(0x15, 0x65, 0xC0)),
            BlokTuru.Ornek => KutuCiz("▸ " + blok.Etiket.ToUpperInvariant(), blok.Metin,
                Color.FromRgb(0xF5, 0xF5, 0xF5), Color.FromRgb(0x61, 0x61, 0x61)),

            _ => new TextBlock { Text = blok.Metin, TextWrapping = TextWrapping.Wrap }
        };

        private UIElement ListeCiz(IReadOnlyList<string> ogeler, bool numarali)
        {
            var panel = new StackPanel { Margin = new Thickness(4, 0, 0, 14) };

            for (int i = 0; i < ogeler.Count; i++)
            {
                var satir = new Grid { Margin = new Thickness(0, 0, 0, 7) };
                satir.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(numarali ? 26 : 18) });
                satir.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                var isaret = new TextBlock
                {
                    Text = numarali ? $"{i + 1}." : "•",
                    FontSize = 13.5,
                    FontWeight = numarali ? FontWeights.Bold : FontWeights.Normal,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x15, 0x65, 0xC0)),
                    VerticalAlignment = VerticalAlignment.Top
                };
                Grid.SetColumn(isaret, 0);

                var metin = new TextBlock
                {
                    Text = ogeler[i],
                    FontSize = 13.5,
                    LineHeight = 21,
                    TextWrapping = TextWrapping.Wrap,
                    Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
                };
                Grid.SetColumn(metin, 1);

                satir.Children.Add(isaret);
                satir.Children.Add(metin);
                panel.Children.Add(satir);
            }

            return panel;
        }

        private UIElement TusCiz(string tus, string aciklama)
        {
            var grid = new Grid();
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
            grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

            var tusKutu = new Border
            {
                Background = new SolidColorBrush(Color.FromRgb(0x26, 0x32, 0x38)),
                CornerRadius = new CornerRadius(5),
                Padding = new Thickness(13, 7, 13, 7),
                MinWidth = 62,
                VerticalAlignment = VerticalAlignment.Center,
                Child = new TextBlock
                {
                    Text = tus,
                    Foreground = Brushes.White,
                    FontWeight = FontWeights.Bold,
                    FontSize = 13,
                    HorizontalAlignment = HorizontalAlignment.Center
                }
            };
            Grid.SetColumn(tusKutu, 0);

            var metin = new TextBlock
            {
                Text = aciklama,
                FontSize = 13.5,
                TextWrapping = TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(14, 0, 0, 0),
                Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
            };
            Grid.SetColumn(metin, 1);

            grid.Children.Add(tusKutu);
            grid.Children.Add(metin);

            return new Border
            {
                Background = Brushes.White,
                BorderBrush = new SolidColorBrush(Color.FromRgb(0xE0, 0xE0, 0xE0)),
                BorderThickness = new Thickness(1),
                CornerRadius = new CornerRadius(5),
                Padding = new Thickness(12, 10, 14, 10),
                Margin = new Thickness(0, 0, 0, 10),
                Child = grid
            };
        }

        private UIElement KutuCiz(string baslik, string metin, Color zemin, Color vurgu)
        {
            var ic = new StackPanel();
            ic.Children.Add(new TextBlock
            {
                Text = baslik,
                FontSize = 10.5,
                FontWeight = FontWeights.Bold,
                Foreground = new SolidColorBrush(vurgu),
                Margin = new Thickness(0, 0, 0, 5)
            });
            ic.Children.Add(new TextBlock
            {
                Text = metin,
                FontSize = 13,
                LineHeight = 21,
                TextWrapping = TextWrapping.Wrap,
                Foreground = new SolidColorBrush(Color.FromRgb(0x33, 0x33, 0x33))
            });

            return new Border
            {
                Background = new SolidColorBrush(zemin),
                BorderBrush = new SolidColorBrush(vurgu),
                BorderThickness = new Thickness(3, 0, 0, 0),
                Padding = new Thickness(14, 11, 14, 12),
                Margin = new Thickness(0, 4, 0, 14),
                Child = ic
            };
        }

        // ── Düğmeler ──────────────────────────────────────────────────────────

        private void Geri_Click(object sender, RoutedEventArgs e) => Goster(_sira - 1);

        private void Ileri_Click(object sender, RoutedEventArgs e)
        {
            if (_sira >= _adimlar.Count - 1)
            {
                IlkCalistirmayiIsaretle();
                Close();
                return;
            }
            Goster(_sira + 1);
        }
    }
}
