using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using UIElementInspector.Core.Models;
using UIElementInspector.Core.Utils;

namespace UIElementInspector.Windows
{
    /// <summary>
    /// Güncelleme penceresi. İki yoldan açılır:
    ///  • Yardım ▸ Güncellemeleri kontrol et  → elle, her zaman sonuç gösterir
    ///  • Açılışta arka plan kontrolü         → yalnız yeni sürüm varsa açılır
    /// </summary>
    public partial class UpdateWindow : Window
    {
        private readonly AppSettings _ayarlar;
        private readonly CancellationTokenSource _iptal = new();
        private GuncellemeSonucu? _sonuc;
        private string? _indirilenDosya;
        private bool _islemSuruyor;

        /// <param name="hazirSonuc">
        /// Arka plan kontrolü zaten yapıldıysa tekrar sorulmasın diye geçilir.
        /// </param>
        public UpdateWindow(AppSettings ayarlar, GuncellemeSonucu? hazirSonuc = null)
        {
            InitializeComponent();
            _ayarlar = ayarlar ?? AppSettings.CreateDefault();
            _sonuc = hazirSonuc;

            txtMevcutSurum.Text = "v" + UpdateService.MevcutSurumMetni;

            Loaded += async (_, __) =>
            {
                if (_sonuc != null) SonucuGoster(_sonuc);
                else await KontrolEtAsync();
            };
        }

        private async Task KontrolEtAsync()
        {
            pnlBekleme.Visibility = Visibility.Visible;
            scrollNotlar.Visibility = Visibility.Collapsed;
            txtBekleme.Text = "GitHub'a bağlanılıyor…";

            var sonuc = await UpdateService.KontrolEtAsync(_ayarlar.GuncellemeDeposu, _iptal.Token);

            _ayarlar.SonGuncellemeKontrolu = DateTime.Now;
            try { _ayarlar.Save(); } catch { /* ayar yazılamazsa kontrol yine de geçerli */ }

            _sonuc = sonuc;
            SonucuGoster(sonuc);
        }

        private void SonucuGoster(GuncellemeSonucu sonuc)
        {
            pnlBekleme.Visibility = Visibility.Collapsed;
            scrollNotlar.Visibility = Visibility.Visible;

            if (!sonuc.Basarili)
            {
                txtAltBaslik.Text = "Kontrol edilemedi";
                txtYeniSurum.Text = "?";
                txtYeniSurum.Foreground = System.Windows.Media.Brushes.Gray;
                txtNotlar.Text =
                    "Güncelleme kontrolü yapılamadı.\n\n" +
                    $"Sebep: {sonuc.Hata}\n\n" +
                    "Bu, yeni sürüm olmadığı anlamına GELMEZ — sadece bakılamadı. " +
                    "İnternet bağlantınızı kontrol edip tekrar deneyebilir ya da " +
                    "aşağıdaki \"GitHub'da aç\" düğmesiyle sürümlere elle bakabilirsiniz.";
                btnGuncelle.Content = "🔄 Tekrar dene";
                btnGuncelle.IsEnabled = true;
                btnGuncelle.Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0));
                return;
            }

            var yeni = sonuc.Yeni!;
            txtYeniSurum.Text = "v" + yeni.Surum.ToString(3);

            if (!sonuc.GuncellemeVar)
            {
                txtAltBaslik.Text = "En güncel sürümü kullanıyorsunuz";
                txtYeniSurum.Foreground = System.Windows.Media.Brushes.Gray;
                txtNotlar.Text =
                    $"Kurulu sürüm: v{UpdateService.MevcutSurumMetni}\n" +
                    $"Yayındaki en son sürüm: {yeni.Etiket}\n\n" +
                    "Güncellemeye gerek yok.\n\n" +
                    (string.IsNullOrWhiteSpace(yeni.Notlar) ? "" : "Son sürüm notları:\n\n" + yeni.Notlar);
                btnGuncelle.Content = "🔄 Tekrar kontrol et";
                btnGuncelle.IsEnabled = true;
                btnGuncelle.Background = new System.Windows.Media.SolidColorBrush(
                    System.Windows.Media.Color.FromRgb(0x15, 0x65, 0xC0));
                return;
            }

            // Yeni sürüm var.
            txtAltBaslik.Text = $"Yeni sürüm hazır — {yeni.Baslik}";
            txtYeniBoyut.Text = yeni.Boyut > 0 ? yeni.BoyutMetni : "";
            btnAtla.Visibility = Visibility.Visible;
            btnSonra.Content = "Sonra";

            txtNotlar.Text = string.IsNullOrWhiteSpace(yeni.Notlar)
                ? "Bu sürüm için not girilmemiş."
                : yeni.Notlar;

            if (sonuc.DosyaEksik)
            {
                txtNotlar.Text +=
                    "\n\n⚠️ Bu sürüme kurulum dosyası eklenmemiş. " +
                    "\"GitHub'da aç\" ile sayfayı açıp elle indirmeniz gerekiyor.";
                btnGuncelle.IsEnabled = false;
                return;
            }

            btnGuncelle.IsEnabled = true;
        }

        private async void Guncelle_Click(object sender, RoutedEventArgs e)
        {
            if (_islemSuruyor) return;

            // Kontrol başarısızsa / güncelleme yoksa düğme "tekrar dene" işlevinde.
            if (_sonuc == null || !_sonuc.GuncellemeVar)
            {
                _sonuc = null;
                await KontrolEtAsync();
                return;
            }

            _islemSuruyor = true;
            btnGuncelle.IsEnabled = false;
            btnAtla.IsEnabled = false;
            btnSonra.Content = "İptal";
            pnlIndirme.Visibility = Visibility.Visible;

            var yeni = _sonuc.Yeni!;
            var ilerleme = new Progress<(long inen, long toplam)>(d =>
            {
                if (d.toplam > 0)
                {
                    var yuzde = d.inen * 100.0 / d.toplam;
                    pbIndirme.IsIndeterminate = false;
                    pbIndirme.Value = yuzde;
                    txtIndirmeYuzde.Text = $"{yuzde:0}%";
                    txtIndirmeDurum.Text =
                        $"İndiriliyor… {d.inen / 1024d / 1024d:0.#} / {d.toplam / 1024d / 1024d:0.#} MB";
                }
                else
                {
                    pbIndirme.IsIndeterminate = true;
                    txtIndirmeDurum.Text = $"İndiriliyor… {d.inen / 1024d / 1024d:0.#} MB";
                    txtIndirmeYuzde.Text = "";
                }
            });

            try
            {
                _indirilenDosya = await UpdateService.IndirAsync(yeni, ilerleme, _iptal.Token);

                txtIndirmeDurum.Text = "İndirme tamam. Kurulum başlatılıyor…";
                txtIndirmeYuzde.Text = "100%";
                pbIndirme.Value = 100;

                var cevap = System.Windows.MessageBox.Show(
                    this,
                    $"Yeni sürüm indirildi (v{yeni.Surum.ToString(3)}).\n\n" +
                    "Kurulum için UIBUL kapatılacak, güncelleme yapılacak ve program " +
                    "yeniden açılacak.\n\nŞimdi devam edilsin mi?",
                    "Güncelleme hazır",
                    MessageBoxButton.OKCancel, MessageBoxImage.Information);

                if (cevap == MessageBoxResult.OK)
                {
                    UpdateService.KurVeCik(_indirilenDosya);
                    return;
                }

                txtIndirmeDurum.Text = "Kurulum ertelendi. İndirilen dosya: " + _indirilenDosya;
                btnGuncelle.Content = "▶ Kurulumu başlat";
                btnGuncelle.IsEnabled = true;
                btnSonra.Content = "Kapat";
                _islemSuruyor = false;
            }
            catch (OperationCanceledException)
            {
                txtIndirmeDurum.Text = "İndirme iptal edildi.";
                pnlIndirme.Visibility = Visibility.Collapsed;
                btnGuncelle.IsEnabled = true;
                btnAtla.IsEnabled = true;
                btnSonra.Content = "Sonra";
                _islemSuruyor = false;
            }
            catch (Exception ex)
            {
                pnlIndirme.Visibility = Visibility.Collapsed;
                System.Windows.MessageBox.Show(this,
                    "Güncelleme indirilemedi.\n\n" + ex.Message +
                    "\n\n\"GitHub'da aç\" ile elle indirebilirsiniz.",
                    "İndirme hatası", MessageBoxButton.OK, MessageBoxImage.Warning);
                btnGuncelle.IsEnabled = true;
                btnAtla.IsEnabled = true;
                btnSonra.Content = "Sonra";
                _islemSuruyor = false;
            }
        }

        private void Atla_Click(object sender, RoutedEventArgs e)
        {
            if (_sonuc?.Yeni != null)
            {
                _ayarlar.AtlananSurum = _sonuc.Yeni.Etiket;
                try { _ayarlar.Save(); } catch { }
            }
            Close();
        }

        private void Sonra_Click(object sender, RoutedEventArgs e)
        {
            if (_islemSuruyor) { _iptal.Cancel(); return; }
            Close();
        }

        private void DepoyuAc_Click(object sender, RoutedEventArgs e)
            => UpdateService.DepoyuAc(_ayarlar.GuncellemeDeposu);

        protected override void OnClosed(EventArgs e)
        {
            try { _iptal.Cancel(); _iptal.Dispose(); } catch { }
            base.OnClosed(e);
        }
    }
}
