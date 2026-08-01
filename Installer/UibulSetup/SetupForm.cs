using System.Diagnostics;

namespace UibulSetup;

/// <summary>
/// Kurulum sihirbazı. Beş sayfa: Hoş geldiniz → Gereklilikler → Seçenekler →
/// Kurulum → Bitti. Arayüz kodla kurulur (tasarımcı dosyası yok) — böylece
/// tek bir dosyada okunabilir kalıyor.
/// </summary>
public sealed class SetupForm : Form
{
    // ── Renkler ──
    private static readonly Color Lacivert = Color.FromArgb(0x15, 0x65, 0xC0);
    private static readonly Color KoyuGri = Color.FromArgb(0x26, 0x32, 0x38);
    private static readonly Color Yesil = Color.FromArgb(0x2E, 0x7D, 0x32);
    private static readonly Color Turuncu = Color.FromArgb(0xE6, 0x51, 0x00);
    private static readonly Color Kirmizi = Color.FromArgb(0xC6, 0x28, 0x28);
    private static readonly Color AcikZemin = Color.FromArgb(0xFA, 0xFA, 0xFA);

    private readonly Label _lblBaslik = new();
    private readonly Label _lblAltBaslik = new();
    private readonly Panel _pnlIcerik = new();
    private readonly Button _btnGeri = new();
    private readonly Button _btnIleri = new();
    private readonly Button _btnIptal = new();

    private readonly KurulumSecenekleri _secenek = new();
    private readonly CancellationTokenSource _iptal = new();

    private List<Gereklilik> _gereklilikler = new();
    private readonly Dictionary<Gereklilik, (Label isaret, Label ayrinti, Button? kur)> _satirlar = new();

    private int _sayfa;
    private bool _kurulumBitti;
    private bool _kurulumSuruyor;

    private const int SonSayfa = 4;

    public SetupForm(string? hedefKlasor)
    {
        if (!string.IsNullOrWhiteSpace(hedefKlasor))
            _secenek.HedefKlasor = hedefKlasor!;

        PencereyiKur();
        Sayfaya(0);
    }

    // ── Pencere iskeleti ──────────────────────────────────────────────────────

    private void PencereyiKur()
    {
        Text = $"{Kurulum.UygulamaAdi} — Kurulum";
        ClientSize = new Size(680, 540);
        MinimumSize = new Size(696, 579);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        StartPosition = FormStartPosition.CenterScreen;
        BackColor = AcikZemin;
        Font = new Font("Segoe UI", 9F);
        AutoScaleMode = AutoScaleMode.Dpi;

        try
        {
            if (Environment.ProcessPath is { } yol)
                Icon = Icon.ExtractAssociatedIcon(yol);
        }
        catch { }

        // Üst başlık şeridi
        var ust = new Panel { Dock = DockStyle.Top, Height = 78, BackColor = Lacivert };

        _lblBaslik.Text = "";
        _lblBaslik.ForeColor = Color.White;
        _lblBaslik.Font = new Font("Segoe UI", 15F, FontStyle.Bold);
        _lblBaslik.AutoSize = false;
        _lblBaslik.Location = new Point(26, 16);
        _lblBaslik.Size = new Size(620, 30);

        _lblAltBaslik.Text = "";
        _lblAltBaslik.ForeColor = Color.FromArgb(0xBB, 0xDE, 0xFB);
        _lblAltBaslik.Font = new Font("Segoe UI", 9F);
        _lblAltBaslik.AutoSize = false;
        _lblAltBaslik.Location = new Point(28, 48);
        _lblAltBaslik.Size = new Size(620, 20);

        ust.Controls.Add(_lblBaslik);
        ust.Controls.Add(_lblAltBaslik);

        // Alt düğme şeridi
        var alt = new Panel { Dock = DockStyle.Bottom, Height = 60, BackColor = Color.FromArgb(0xF0, 0xF0, 0xF0) };

        DugmeBicimle(_btnIptal, "İptal", 150);
        _btnIptal.Location = new Point(24, 14);
        _btnIptal.Click += (_, _) => IptalEt();

        DugmeBicimle(_btnGeri, "◀ Geri", 105);
        _btnGeri.Location = new Point(414, 14);
        _btnGeri.Click += (_, _) => Sayfaya(_sayfa - 1);

        DugmeBicimle(_btnIleri, "İleri ▶", 135, birincil: true);
        _btnIleri.Location = new Point(527, 14);
        _btnIleri.Click += async (_, _) => await IleriBas();

        alt.Controls.AddRange(new Control[] { _btnIptal, _btnGeri, _btnIleri });

        _pnlIcerik.Dock = DockStyle.Fill;
        _pnlIcerik.BackColor = AcikZemin;
        _pnlIcerik.Padding = new Padding(26, 20, 26, 12);
        _pnlIcerik.AutoScroll = true;

        Controls.Add(_pnlIcerik);
        Controls.Add(alt);
        Controls.Add(ust);

        FormClosing += (_, e) =>
        {
            if (_kurulumSuruyor)
            {
                var c = MessageBox.Show(this,
                    "Kurulum sürüyor. Yarıda kesilirse program eksik kurulmuş olabilir.\n\nYine de çıkılsın mı?",
                    "Kurulum sürüyor", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (c != DialogResult.Yes) { e.Cancel = true; return; }
                _iptal.Cancel();
            }
        };
    }

    private static void DugmeBicimle(Button d, string metin, int genislik, bool birincil = false)
    {
        d.Text = metin;
        d.Size = new Size(genislik, 33);
        d.FlatStyle = FlatStyle.Flat;
        d.Font = new Font("Segoe UI", 9.5F, birincil ? FontStyle.Bold : FontStyle.Regular);
        d.Cursor = Cursors.Hand;

        if (birincil)
        {
            d.BackColor = Lacivert;
            d.ForeColor = Color.White;
            d.FlatAppearance.BorderSize = 0;
        }
        else
        {
            d.BackColor = Color.White;
            d.ForeColor = Color.FromArgb(0x55, 0x55, 0x55);
            d.FlatAppearance.BorderColor = Color.FromArgb(0xBD, 0xBD, 0xBD);
        }
    }

    // ── Sayfa yönetimi ────────────────────────────────────────────────────────

    private void Sayfaya(int sayfa)
    {
        if (sayfa < 0 || sayfa > SonSayfa) return;
        _sayfa = sayfa;

        _pnlIcerik.Controls.Clear();
        _pnlIcerik.AutoScrollPosition = Point.Empty;

        _btnGeri.Visible = sayfa is > 0 and < 3;
        _btnIptal.Visible = sayfa < 3;
        _btnIleri.Enabled = true;

        switch (sayfa)
        {
            case 0: SayfaHosgeldiniz(); break;
            case 1: SayfaGereklilikler(); break;
            case 2: SayfaSecenekler(); break;
            case 3: SayfaKurulum(); break;
            case 4: SayfaBitti(); break;
        }
    }

    private async Task IleriBas()
    {
        switch (_sayfa)
        {
            case 0:
                Sayfaya(1);
                break;

            case 1:
                Sayfaya(2);
                break;

            case 2:
                if (!HedefiDogrula()) return;

                // Program Files gibi korumalı bir yer seçildiyse kendini yükselt.
                if (Kurulum.YoneticiGerekliMi(_secenek.HedefKlasor))
                {
                    var c = MessageBox.Show(this,
                        $"Seçtiğiniz klasöre yazmak için yönetici hakkı gerekiyor:\n{_secenek.HedefKlasor}\n\n" +
                        "Kurulum yönetici olarak yeniden başlatılsın mı?\n\n" +
                        "(\"Hayır\" derseniz kendi kullanıcı klasörünüze kurabilirsiniz — " +
                        "önerilen yol budur, güncellemeler de yönetici sormadan çalışır.)",
                        "Yönetici hakkı gerekiyor",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                    if (c == DialogResult.Yes)
                    {
                        if (Kurulum.KendiniYukselt($"/KLASOR:\"{_secenek.HedefKlasor}\""))
                        {
                            Application.Exit();
                            return;
                        }
                        MessageBox.Show(this,
                            "Yönetici olarak başlatılamadı. Lütfen başka bir klasör seçin.",
                            "Yükseltme başarısız", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                    return;
                }

                Sayfaya(3);
                await KurulumuCalistir();
                break;

            case 3:
                if (_kurulumBitti) Sayfaya(4);
                break;

            case 4:
                Bitir();
                break;
        }
    }

    private void IptalEt()
    {
        var c = MessageBox.Show(this, "Kurulumdan çıkılsın mı?", "Kurulum",
            MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (c == DialogResult.Yes) { _iptal.Cancel(); Application.Exit(); }
    }

    private void Bitir()
    {
        if (_secenek.KurulumSonrasiCalistir)
        {
            try
            {
                var exe = Path.Combine(_secenek.HedefKlasor, Kurulum.ExeAdi);
                if (File.Exists(exe))
                    Process.Start(new ProcessStartInfo
                    {
                        FileName = exe,
                        WorkingDirectory = _secenek.HedefKlasor,
                        UseShellExecute = true
                    });
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "Program başlatılamadı: " + ex.Message,
                    "UIBUL", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        Application.Exit();
    }

    // ── Sayfa 0: Hoş geldiniz ─────────────────────────────────────────────────

    private void SayfaHosgeldiniz()
    {
        Baslik("Hoş geldiniz", $"{Kurulum.UygulamaAdi} sürüm {Kurulum.SetupSurumu} kurulacak");

        var y = 8;

        y = MetinEkle(y,
            "Bu sihirbaz UIBUL'u bilgisayarınıza kuracak. Kurulum birkaç dakika sürer ve " +
            "başka hiçbir program indirmenizi gerektirmez — ihtiyaç duyulan her şey bu " +
            "dosyanın içindedir.", 11F);

        y += 6;
        y = BaslikEkle(y, "Kurulum ne yapacak?");
        y = MaddeEkle(y, "Uygulamayı kendi kullanıcı klasörünüze kuracak (yönetici hakkı gerekmez)");
        y = MaddeEkle(y, "Bilgisayarınızın gerekliliklerini denetleyecek, eksik varsa kuracak");
        y = MaddeEkle(y, "Masaüstü ve Başlat menüsü kısayolları oluşturacak");
        y = MaddeEkle(y, "Programı Windows'un \"Uygulamalar\" listesine kaydedecek");
        y = MaddeEkle(y, "Otomatik güncelleme sistemini etkinleştirecek");

        y += 10;

        var mevcut = Kurulum.MevcutKurulumKlasoru();
        if (mevcut != null)
        {
            var eski = Kurulum.KuruluSurum() ?? "bilinmeyen";
            _secenek.HedefKlasor = mevcut;
            KutuEkle(ref y, "ℹ️ MEVCUT KURULUM BULUNDU",
                $"UIBUL zaten kurulu (sürüm {eski}).\n" +
                $"Klasör: {mevcut}\n\n" +
                $"Devam ederseniz sürüm {Kurulum.SetupSurumu} olarak güncellenecek. " +
                "Ayarlarınız, arşiviniz ve ekran görüntüleriniz korunur.",
                Lacivert, Color.FromArgb(0xE3, 0xF2, 0xFD));
        }
        else
        {
            KutuEkle(ref y, "💡 İLK KURULUM",
                "Kurulum bittikten sonra program ilk açılışında size adım adım bir " +
                "öğretici gösterecek. Hangi tuşun ne yaptığını, günlük işlerde nasıl " +
                "kullanabileceğinizi orada bulacaksınız.",
                Yesil, Color.FromArgb(0xE8, 0xF5, 0xE9));
        }

        _btnIleri.Text = "İleri ▶";
    }

    // ── Sayfa 1: Gereklilikler ────────────────────────────────────────────────

    private void SayfaGereklilikler()
    {
        Baslik("Sistem denetimi", "Bilgisayarınızın UIBUL'u çalıştırabildiği doğrulanıyor");

        _satirlar.Clear();
        var gerekliBayt = Math.Max(Kurulum.YukBoyutu(), 200L * 1024 * 1024);
        _gereklilikler = Gereklilikler.Olustur(_secenek.HedefKlasor, gerekliBayt);

        var y = 4;
        foreach (var g in _gereklilikler)
        {
            var satir = new Panel
            {
                Location = new Point(0, y),
                Size = new Size(600, 52),
                BackColor = Color.White,
                BorderStyle = BorderStyle.FixedSingle
            };

            var isaret = new Label
            {
                Text = "○",
                Font = new Font("Segoe UI", 13F, FontStyle.Bold),
                ForeColor = Color.Gray,
                Location = new Point(12, 13),
                Size = new Size(28, 26),
                TextAlign = ContentAlignment.MiddleCenter
            };

            var ad = new Label
            {
                Text = g.Ad + (g.Zorunlu ? "" : "  (isteğe bağlı)"),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Location = new Point(46, 8),
                Size = new Size(400, 18),
                ForeColor = Color.FromArgb(0x21, 0x21, 0x21)
            };

            var ayrinti = new Label
            {
                Text = g.Aciklama,
                Font = new Font("Segoe UI", 8.5F),
                ForeColor = Color.FromArgb(0x75, 0x75, 0x75),
                Location = new Point(46, 27),
                Size = new Size(420, 18)
            };

            Button? kurDugmesi = null;
            if (g.Kur != null)
            {
                kurDugmesi = new Button { Visible = false };
                DugmeBicimle(kurDugmesi, "Kur", 88);
                kurDugmesi.Size = new Size(88, 27);
                kurDugmesi.Location = new Point(492, 12);
                var yerel = g;
                kurDugmesi.Click += async (_, _) => await EksigiKur(yerel);
                satir.Controls.Add(kurDugmesi);
            }

            satir.Controls.AddRange(new Control[] { isaret, ad, ayrinti });
            _pnlIcerik.Controls.Add(satir);
            _satirlar[g] = (isaret, ayrinti, kurDugmesi);

            y += 58;
        }

        _btnIleri.Text = "İleri ▶";
        _btnIleri.Enabled = false;

        _ = DenetimleriCalistir();
    }

    private async Task DenetimleriCalistir()
    {
        foreach (var g in _gereklilikler)
        {
            g.Sonuc = Durum.Deneniyor;
            SatiriGuncelle(g, "denetleniyor…");
            await Task.Delay(90);   // kullanıcı sırayla denetlendiğini görsün

            try
            {
                var (tamam, ayrinti) = g.Denetle();
                g.Ayrinti = ayrinti;
                g.Sonuc = tamam ? Durum.Tamam
                        : g.Kur != null ? Durum.Kurulacak
                        : g.Zorunlu ? Durum.Eksik
                        : Durum.Atlandi;
            }
            catch (Exception ex)
            {
                g.Ayrinti = "denetlenemedi: " + ex.Message;
                g.Sonuc = g.Zorunlu ? Durum.Eksik : Durum.Atlandi;
            }

            SatiriGuncelle(g, g.Ayrinti);
        }

        SonucuOzetle();
    }

    private void SatiriGuncelle(Gereklilik g, string ayrinti)
    {
        if (!_satirlar.TryGetValue(g, out var s)) return;

        (s.isaret.Text, s.isaret.ForeColor) = g.Sonuc switch
        {
            Durum.Deneniyor => ("◌", Color.FromArgb(0x90, 0xA4, 0xAE)),
            Durum.Tamam => ("✔", Yesil),
            Durum.Kurulacak => ("!", Turuncu),
            Durum.Eksik => ("✖", Kirmizi),
            Durum.Atlandi => ("–", Color.FromArgb(0x90, 0xA4, 0xAE)),
            _ => ("○", Color.Gray)
        };

        s.ayrinti.Text = ayrinti;
        s.ayrinti.ForeColor = g.Sonuc == Durum.Eksik ? Kirmizi
                            : g.Sonuc == Durum.Kurulacak ? Turuncu
                            : Color.FromArgb(0x75, 0x75, 0x75);

        if (s.kur != null) s.kur.Visible = g.Sonuc == Durum.Kurulacak;

        Application.DoEvents();
    }

    private async Task EksigiKur(Gereklilik g)
    {
        if (g.Kur == null) return;

        if (_satirlar.TryGetValue(g, out var s) && s.kur != null)
        {
            s.kur.Enabled = false;
            s.kur.Text = "…";
        }

        g.Sonuc = Durum.Deneniyor;
        SatiriGuncelle(g, "kuruluyor…");

        var basarili = await g.Kur(m => SatiriGuncelle(g, m), _iptal.Token);

        var (tamam, ayrinti) = g.Denetle();
        g.Ayrinti = ayrinti;
        g.Sonuc = tamam ? Durum.Tamam : (g.Zorunlu ? Durum.Eksik : Durum.Atlandi);
        SatiriGuncelle(g, basarili || tamam ? ayrinti : ayrinti + " — kurulamadı");

        if (_satirlar.TryGetValue(g, out var s2) && s2.kur != null)
        {
            s2.kur.Enabled = true;
            s2.kur.Text = "Kur";
        }

        SonucuOzetle();
    }

    private void SonucuOzetle()
    {
        var engel = _gereklilikler.Any(g => g.Zorunlu && g.Sonuc == Durum.Eksik);
        _btnIleri.Enabled = !engel;

        var eksikIstege = _gereklilikler.Count(g => !g.Zorunlu && g.Sonuc is Durum.Kurulacak or Durum.Atlandi);

        _lblAltBaslik.Text = engel
            ? "Zorunlu bir gereklilik karşılanmıyor — kuruluma devam edilemiyor"
            : eksikIstege > 0
                ? $"Her şey hazır. {eksikIstege} isteğe bağlı bileşen eksik (kurulum yine de yapılabilir)"
                : "Her şey hazır, kuruluma geçebilirsiniz";
    }

    // ── Sayfa 2: Seçenekler ───────────────────────────────────────────────────

    private void SayfaSecenekler()
    {
        Baslik("Kurulum seçenekleri", "Nereye kurulacağını ve kısayolları seçin");

        var y = 4;
        y = BaslikEkle(y, "Kurulum klasörü");

        var kutu = new TextBox
        {
            Text = _secenek.HedefKlasor,
            Location = new Point(2, y),
            Size = new Size(500, 26),
            Font = new Font("Segoe UI", 9.5F)
        };
        kutu.TextChanged += (_, _) => _secenek.HedefKlasor = kutu.Text.Trim();

        var gozat = new Button();
        DugmeBicimle(gozat, "Gözat…", 92);
        gozat.Size = new Size(92, 27);
        gozat.Location = new Point(510, y - 1);
        gozat.Click += (_, _) =>
        {
            using var d = new FolderBrowserDialog
            {
                Description = "UIBUL'un kurulacağı klasörü seçin",
                UseDescriptionForTitle = true,
                SelectedPath = Directory.Exists(_secenek.HedefKlasor)
                    ? _secenek.HedefKlasor
                    : Kurulum.VarsayilanKlasor()
            };
            if (d.ShowDialog(this) == DialogResult.OK)
                kutu.Text = Path.Combine(d.SelectedPath,
                    Path.GetFileName(d.SelectedPath).Equals(Kurulum.KisaAd, StringComparison.OrdinalIgnoreCase)
                        ? "" : Kurulum.KisaAd);
        };

        _pnlIcerik.Controls.Add(kutu);
        _pnlIcerik.Controls.Add(gozat);
        y += 34;

        var boyut = Kurulum.YukBoyutu();
        y = NotEkle(y, boyut > 0
            ? $"Gereken alan: yaklaşık {boyut / 1024d / 1024:0} MB"
            : "Gereken alan hesaplanamadı");

        y += 14;
        y = BaslikEkle(y, "Kısayollar ve başlangıç");

        y = OnayEkle(y, "Masaüstüne kısayol oluştur", _secenek.MasaustuKisayolu,
            v => _secenek.MasaustuKisayolu = v);
        y = OnayEkle(y, "Başlat menüsüne ekle", _secenek.BaslatMenusuKisayolu,
            v => _secenek.BaslatMenusuKisayolu = v);
        y = OnayEkle(y, "Windows açılışında otomatik başlat", _secenek.WindowsIleBaslat,
            v => _secenek.WindowsIleBaslat = v);
        y = NotEkle(y, "Kısayol tuşları (F1–F11) yalnızca program açıkken çalışır. " +
                       "Ekran görüntüsü için sürekli hazır olmasını istiyorsanız bunu işaretleyin.");

        y += 14;
        KutuEkle(ref y, "⚠️ KISAYOL TUŞLARI HAKKINDA",
            "UIBUL açıkken F1–F11 tuşlarını sistem genelinde kendine alır. Bu, tuşların " +
            "her uygulamada çalışmasını sağlar; ama o tuşları kullanan başka programlar " +
            "(örneğin Chrome'un F11 tam ekranı) UIBUL açıkken tepki vermez. " +
            "Rahatsız ederse ayarlardan değiştirebilir ya da programı kapatabilirsiniz.",
            Turuncu, Color.FromArgb(0xFF, 0xF3, 0xE0));

        _btnIleri.Text = "Kur ▶";
    }

    private bool HedefiDogrula()
    {
        var yol = _secenek.HedefKlasor?.Trim() ?? "";

        if (string.IsNullOrWhiteSpace(yol))
        {
            MessageBox.Show(this, "Kurulum klasörü boş olamaz.", "Kurulum klasörü",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        try
        {
            yol = Path.GetFullPath(yol);
            if (!Path.IsPathRooted(yol)) throw new ArgumentException();
            _secenek.HedefKlasor = yol;
            return true;
        }
        catch
        {
            MessageBox.Show(this, "Klasör yolu geçersiz:\n" + yol, "Kurulum klasörü",
                MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
    }

    // ── Sayfa 3: Kurulum ──────────────────────────────────────────────────────

    private ProgressBar _cubuk = new();
    private Label _lblDurum = new();
    private ListBox _gunluk = new();

    private void SayfaKurulum()
    {
        Baslik("Kuruluyor", "Lütfen bekleyin, bu birkaç dakika sürebilir");

        _btnIleri.Text = "İleri ▶";
        _btnIleri.Enabled = false;

        _lblDurum = new Label
        {
            Text = "Hazırlanıyor…",
            Location = new Point(2, 8),
            Size = new Size(600, 20),
            Font = new Font("Segoe UI", 9.5F)
        };

        _cubuk = new ProgressBar
        {
            Location = new Point(2, 34),
            Size = new Size(600, 18),
            Minimum = 0,
            Maximum = 100,
            Style = ProgressBarStyle.Continuous
        };

        _gunluk = new ListBox
        {
            Location = new Point(2, 66),
            Size = new Size(600, 300),
            Font = new Font("Consolas", 8.5F),
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = Color.White,
            ForeColor = Color.FromArgb(0x42, 0x42, 0x42)
        };

        _pnlIcerik.Controls.AddRange(new Control[] { _lblDurum, _cubuk, _gunluk });
    }

    private async Task KurulumuCalistir()
    {
        _kurulumSuruyor = true;

        void Bildir(string metin, int yuzde)
        {
            if (InvokeRequired) { BeginInvoke(() => Bildir(metin, yuzde)); return; }

            _lblDurum.Text = metin;
            if (yuzde >= 0) _cubuk.Value = Math.Clamp(yuzde, 0, 100);

            if (_gunluk.Items.Count == 0 || (string)_gunluk.Items[^1]! != metin)
            {
                _gunluk.Items.Add(metin);
                _gunluk.TopIndex = _gunluk.Items.Count - 1;
            }
        }

        try
        {
            Bildir($"Hedef: {_secenek.HedefKlasor}", 0);
            await Task.Run(() => Kurulum.Kur(_secenek, Bildir, _iptal.Token), _iptal.Token);

            _kurulumBitti = true;
            _kurulumSuruyor = false;
            _btnIleri.Enabled = true;
            _btnIptal.Visible = false;

            Bildir("Kurulum başarıyla tamamlandı.", 100);
            Sayfaya(4);
        }
        catch (OperationCanceledException)
        {
            _kurulumSuruyor = false;
            Bildir("Kurulum iptal edildi.", -1);
            _btnIptal.Visible = true;
        }
        catch (Exception ex)
        {
            _kurulumSuruyor = false;
            Bildir("HATA: " + ex.Message, -1);
            _btnIptal.Visible = true;

            MessageBox.Show(this,
                "Kurulum tamamlanamadı.\n\n" + ex.Message +
                "\n\nAntivirüs programınız engelliyor olabilir; kapatıp tekrar deneyin. " +
                "Sorun sürerse başka bir kurulum klasörü seçin.",
                "Kurulum hatası", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // ── Sayfa 4: Bitti ────────────────────────────────────────────────────────

    private void SayfaBitti()
    {
        Baslik("Kurulum tamamlandı", $"{Kurulum.UygulamaAdi} kullanıma hazır");

        _btnIleri.Text = "Bitir ✓";
        _btnIleri.Enabled = true;
        _btnGeri.Visible = false;
        _btnIptal.Visible = false;

        var y = 6;
        y = MetinEkle(y, $"UIBUL sürüm {Kurulum.SetupSurumu} şu klasöre kuruldu:", 10F);
        y = NotEkle(y, _secenek.HedefKlasor);
        y += 8;

        y = BaslikEkle(y, "Şimdi ne yapmalı?");
        y = MaddeEkle(y, "Program ilk açılışta adım adım bir öğretici gösterecek — okuyun, 5 dakika sürer");
        y = MaddeEkle(y, "En çok kullanacağınız tuş F9: ekranın bir bölgesinin görüntüsünü alır");
        y = MaddeEkle(y, "Element incelemeye başlamak için F1, durdurmak için F2");
        y = MaddeEkle(y, "Tanıtım belgesi Başlat menüsündeki UIBUL klasöründe");

        y += 10;
        KutuEkle(ref y, "🔄 GÜNCELLEMELER",
            "UIBUL yeni sürümleri kendisi kontrol eder. Yeni bir sürüm yayınlandığında " +
            "programı açtığınızda haber verilir ve tek tıkla güncellenir — bu kurulumu " +
            "tekrar çalıştırmanıza gerek kalmaz.",
            Lacivert, Color.FromArgb(0xE3, 0xF2, 0xFD));

        y += 6;
        OnayEkle(y, "Programı şimdi başlat", _secenek.KurulumSonrasiCalistir,
            v => _secenek.KurulumSonrasiCalistir = v);
    }

    // ── Küçük arayüz yardımcıları ─────────────────────────────────────────────

    private void Baslik(string baslik, string alt)
    {
        _lblBaslik.Text = baslik;
        _lblAltBaslik.Text = alt;
    }

    private int MetinEkle(int y, string metin, float boyut)
    {
        var l = new Label
        {
            Text = metin,
            Location = new Point(2, y),
            Size = new Size(608, 0),
            AutoSize = false,
            MaximumSize = new Size(608, 0),
            Font = new Font("Segoe UI", boyut),
            ForeColor = Color.FromArgb(0x33, 0x33, 0x33)
        };
        l.Height = TextRenderer.MeasureText(metin, l.Font, new Size(608, 0), TextFormatFlags.WordBreak).Height + 6;
        _pnlIcerik.Controls.Add(l);
        return y + l.Height + 6;
    }

    private int BaslikEkle(int y, string metin)
    {
        var l = new Label
        {
            Text = metin,
            Location = new Point(2, y),
            Size = new Size(608, 22),
            Font = new Font("Segoe UI", 10.5F, FontStyle.Bold),
            ForeColor = Lacivert
        };
        _pnlIcerik.Controls.Add(l);
        return y + 27;
    }

    private int MaddeEkle(int y, string metin)
    {
        var isaret = new Label
        {
            Text = "•",
            Location = new Point(6, y),
            Size = new Size(14, 20),
            Font = new Font("Segoe UI", 10F, FontStyle.Bold),
            ForeColor = Lacivert
        };

        var l = new Label
        {
            Text = metin,
            Location = new Point(22, y),
            AutoSize = false,
            MaximumSize = new Size(586, 0),
            Size = new Size(586, 0),
            Font = new Font("Segoe UI", 9.5F),
            ForeColor = Color.FromArgb(0x33, 0x33, 0x33)
        };
        l.Height = TextRenderer.MeasureText(metin, l.Font, new Size(586, 0), TextFormatFlags.WordBreak).Height + 4;

        _pnlIcerik.Controls.Add(isaret);
        _pnlIcerik.Controls.Add(l);
        return y + Math.Max(22, l.Height + 4);
    }

    private int NotEkle(int y, string metin)
    {
        var l = new Label
        {
            Text = metin,
            Location = new Point(4, y),
            AutoSize = false,
            MaximumSize = new Size(600, 0),
            Size = new Size(600, 0),
            Font = new Font("Segoe UI", 8.5F, FontStyle.Italic),
            ForeColor = Color.FromArgb(0x75, 0x75, 0x75)
        };
        l.Height = TextRenderer.MeasureText(metin, l.Font, new Size(600, 0), TextFormatFlags.WordBreak).Height + 4;
        _pnlIcerik.Controls.Add(l);
        return y + l.Height + 6;
    }

    private int OnayEkle(int y, string metin, bool isaretli, Action<bool> degisti)
    {
        var o = new CheckBox
        {
            Text = metin,
            Checked = isaretli,
            Location = new Point(4, y),
            Size = new Size(600, 24),
            Font = new Font("Segoe UI", 9.5F),
            ForeColor = Color.FromArgb(0x33, 0x33, 0x33)
        };
        o.CheckedChanged += (_, _) => degisti(o.Checked);
        _pnlIcerik.Controls.Add(o);
        return y + 28;
    }

    private void KutuEkle(ref int y, string baslik, string metin, Color vurgu, Color zemin)
    {
        var basl = new Label
        {
            Text = baslik,
            Font = new Font("Segoe UI", 8.5F, FontStyle.Bold),
            ForeColor = vurgu,
            Location = new Point(12, 10),
            Size = new Size(560, 16)
        };

        var govde = new Label
        {
            Text = metin,
            Font = new Font("Segoe UI", 9F),
            ForeColor = Color.FromArgb(0x33, 0x33, 0x33),
            Location = new Point(12, 30),
            AutoSize = false,
            MaximumSize = new Size(566, 0),
            Size = new Size(566, 0)
        };
        govde.Height = TextRenderer.MeasureText(metin, govde.Font, new Size(566, 0),
            TextFormatFlags.WordBreak).Height + 4;

        var kutu = new Panel
        {
            Location = new Point(2, y),
            Size = new Size(600, govde.Height + 44),
            BackColor = zemin,
            Padding = new Padding(3, 0, 0, 0)
        };

        var seritKutu = new Panel
        {
            Dock = DockStyle.Left,
            Width = 4,
            BackColor = vurgu
        };

        kutu.Controls.Add(basl);
        kutu.Controls.Add(govde);
        kutu.Controls.Add(seritKutu);
        _pnlIcerik.Controls.Add(kutu);

        y += kutu.Height + 12;
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing) { try { _iptal.Cancel(); _iptal.Dispose(); } catch { } }
        base.Dispose(disposing);
    }
}
