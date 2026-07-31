using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.Drawing.Text;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Tarih Atlası kusur bildirimi için kare üretir:
    /// kırp → küçült → kırmızı çerçeve → kimlik şeridi → PNG.
    ///
    /// ═══ TASARIMI BELİRLEYEN ÖLÇÜLMÜŞ GERÇEKLER ═══
    ///
    /// 1. Sohbete yapıştırılan görüntünün maliyeti YALNIZCA piksel sayısına
    ///    bağlıdır. Format (PNG/JPEG) ve sıkıştırma kalitesi hiçbir şey
    ///    değiştirmez — görüntü karşı tarafa ulaşmadan piksellere açılır.
    ///    ⇒ Kalite ayarı YOKTUR; olsaydı kullanıcıyı boşuna oyalardı.
    ///    ⇒ Tek gerçek tasarruf küçültmedir (ve bölgeyi dar seçmektir).
    ///
    /// 2. PNG kullanılır, JPEG değil: token farkı yok ama JPEG ince harita
    ///    sınırlarını ve küçük şehir etiketlerini bulanıklaştırır. Bedava kayıp.
    ///
    /// 3. KÜÇÜKSE BÜYÜTÜLMEZ. Büyütmek bulanıklaştırır ve maliyeti artırır;
    ///    tek kazancı yoktur.
    ///
    /// ═══ ADIM SIRASI BAĞLAYICIDIR ═══
    /// Küçültme ÖNCE, çerçeve ve şerit SONRA çizilir. Tersi yapılırsa 3 px'lik
    /// çerçeve 2 px'e iner ve şerit yazısı okunmaz hâle gelir.
    /// </summary>
    public static class AtlasKare
    {
        /// <summary>Kırmızı çerçeve kalınlığı (px) — kırpma sınırı belli olsun diye.</summary>
        public const int CERCEVE = 3;

        private static readonly Color CERCEVE_RENK = Color.FromArgb(255, 220, 20, 30);
        private static readonly Color SERIT_ZEMIN = Color.FromArgb(255, 24, 24, 26);
        // Damga okunamadığında şerit koyu kırmızı olur: karenin tarihsiz
        // olduğu bir bakışta anlaşılsın, sonradan "hangi yıldı bu" diye
        // sorulmasın.
        private static readonly Color SERIT_ZEMIN_UYARI = Color.FromArgb(255, 92, 14, 18);
        private static readonly Color SERIT_YAZI = Color.FromArgb(255, 245, 245, 245);

        private const int SERIT_ASGARI_YUKSEKLIK = 20;
        private const int SERIT_YAN_BOSLUK = 8;

        // Dosya adı deseni: 1361-02-01_41.35N_26.50E_z6.png
        // Budama YALNIZ bu desene uyan dosyalara dokunur; klasördeki başka
        // hiçbir şey silinmez.
        private static readonly Regex ReKareAdi = new Regex(
            @"^(-?\d{1,5}-\d{2}-\d{2}_\d+\.\d+[NS]_\d+\.\d+[EW]_z[\d.]+|damgasiz_\d{4}-\d{2}-\d{2}_\d{6})(\-\d+)?\.png$",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public sealed class Sonuc
        {
            public string DosyaYolu { get; set; }
            public int KaynakGenislik { get; set; }
            public int KaynakYukseklik { get; set; }
            public int SonGenislik { get; set; }
            public int SonYukseklik { get; set; }
            public bool Kucultuldu { get; set; }
            public Bitmap Kare { get; set; }   // pano için — çağıran dispose eder
        }

        /// <summary>
        /// Ekrandan bölgeyi alır, işler, PNG olarak kaydeder.
        /// Panoya kopyalamaz — onu çağıran yapar (Sonuc.Kare hazır durur).
        /// </summary>
        public static Sonuc Uret(Rectangle bolge, string damgaSatiri, string maddeSatiri,
                                 string klasor, string dosyaAdi, int enUzunKenar,
                                 bool uyari = false)
        {
            if (bolge.Width < 1 || bolge.Height < 1)
                throw new ArgumentException("Bölge boş.");

            using (var ham = ScreenshotHelper.CaptureRegion(bolge))
            {
                if (ham == null) throw new Exception("Ekran görüntüsü alınamadı.");

                var sonuc = new Sonuc
                {
                    KaynakGenislik = ham.Width,
                    KaynakYukseklik = ham.Height
                };

                // ── 1. ADIM: küçültme (yalnız küçültme; büyütme YOK) ──────────
                Bitmap icerik = null;
                try
                {
                    int uzunKenar = Math.Max(ham.Width, ham.Height);
                    if (enUzunKenar > 0 && uzunKenar > enUzunKenar)
                    {
                        double oran = (double)enUzunKenar / uzunKenar;
                        int g = Math.Max(1, (int)Math.Round(ham.Width * oran));
                        int y = Math.Max(1, (int)Math.Round(ham.Height * oran));
                        icerik = Kucult(ham, g, y);
                        sonuc.Kucultuldu = true;
                    }
                    else
                    {
                        icerik = new Bitmap(ham);   // kopya: ham using'den çıkacak
                    }

                    // ── 2. ADIM: çerçeve + şerit ──────────────────────────────
                    var kare = Cerceve(icerik, damgaSatiri, maddeSatiri, uyari);
                    sonuc.Kare = kare;
                    sonuc.SonGenislik = kare.Width;
                    sonuc.SonYukseklik = kare.Height;
                }
                finally
                {
                    if (icerik != null) icerik.Dispose();
                }

                // ── 3. ADIM: kaydet ───────────────────────────────────────────
                if (!Directory.Exists(klasor)) Directory.CreateDirectory(klasor);
                var yol = BenzersizYol(Path.Combine(klasor, dosyaAdi));
                sonuc.Kare.Save(yol, ImageFormat.Png);
                sonuc.DosyaYolu = yol;

                return sonuc;
            }
        }

        /// <summary>
        /// Kaliteli küçültme. System.Drawing'de gerçek Lanczos yoktur; bu
        /// ölçekte (0.5–1.0×) HighQualityBicubic ondan ayırt edilemez.
        /// WrapMode.TileFlipXY olmazsa kenarlarda yarı saydam bir hâle çıkar.
        /// </summary>
        private static Bitmap Kucult(Bitmap kaynak, int genislik, int yukseklik)
        {
            var hedef = new Bitmap(genislik, yukseklik, PixelFormat.Format32bppArgb);
            hedef.SetResolution(kaynak.HorizontalResolution, kaynak.VerticalResolution);

            using (var g = Graphics.FromImage(hedef))
            {
                g.CompositingMode = CompositingMode.SourceCopy;
                g.CompositingQuality = CompositingQuality.HighQuality;
                g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                g.SmoothingMode = SmoothingMode.HighQuality;
                g.PixelOffsetMode = PixelOffsetMode.HighQuality;

                using (var oz = new ImageAttributes())
                {
                    oz.SetWrapMode(WrapMode.TileFlipXY);
                    g.DrawImage(kaynak, new Rectangle(0, 0, genislik, yukseklik),
                                0, 0, kaynak.Width, kaynak.Height, GraphicsUnit.Pixel, oz);
                }
            }
            return hedef;
        }

        /// <summary>
        /// İçeriğin çevresine kırmızı çerçeve, altına kimlik şeridi koyar.
        /// Çerçeve içeriğin ÜSTÜNE değil DIŞINA çizilir — 3 px harita örtülmez.
        ///
        /// Şerit bir ya da iki satırdır:
        ///   1) 1361-02-01 · 41.35N 26.50E · z6 · Osmanlı Tarih Atlası
        ///   2) Madde: Edirne'nin fethi              (kronoloji maddesi varsa)
        ///
        /// Birinci satır ASLA kırpılmaz — sığmazsa yazı küçültülür, o da
        /// yetmezse tuval genişletilir. İkinci satır en çok iki satıra sarar,
        /// taşarsa "…" ile biter; madde metinleri sınırsız uzunlukta olabilir
        /// ve tuvali onun için genişletmek kareyi (dolayısıyla maliyeti)
        /// gereksiz büyütürdü.
        /// </summary>
        private static Bitmap Cerceve(Bitmap icerik, string damgaSatiri, string maddeSatiri,
                                      bool uyari)
        {
            int cerceveliG = icerik.Width + 2 * CERCEVE;
            int cerceveliY = icerik.Height + 2 * CERCEVE;
            bool maddeVar = !string.IsNullOrWhiteSpace(maddeSatiri);

            Font font1 = null, font2 = null;
            try
            {
                SizeF olcu1;
                int tuvalG, satir1Yukseklik, maddeYukseklik = 0;

                using (var olcumBmp = new Bitmap(1, 1))
                using (var olcumG = Graphics.FromImage(olcumBmp))
                {
                    font1 = SeritFontu(olcumG, damgaSatiri,
                                       cerceveliG - 2 * SERIT_YAN_BOSLUK, out olcu1);
                    satir1Yukseklik = (int)Math.Ceiling(olcu1.Height);

                    tuvalG = Math.Max(cerceveliG, (int)Math.Ceiling(olcu1.Width) + 2 * SERIT_YAN_BOSLUK);

                    if (maddeVar)
                    {
                        font2 = new Font("Segoe UI", Math.Max(7.5f, font1.SizeInPoints - 0.5f),
                                         FontStyle.Regular, GraphicsUnit.Point);
                        int yerlesimG = tuvalG - 2 * SERIT_YAN_BOSLUK;
                        int enCokIkiSatir = (int)Math.Ceiling(font2.GetHeight(olcumG) * 2);
                        using (var olcumBicim = new StringFormat())
                        {
                            olcumBicim.Trimming = StringTrimming.EllipsisCharacter;
                            var o2 = olcumG.MeasureString(maddeSatiri, font2,
                                        new SizeF(yerlesimG, enCokIkiSatir), olcumBicim);
                            maddeYukseklik = Math.Min(enCokIkiSatir, (int)Math.Ceiling(o2.Height));
                        }
                    }
                }

                int seritYukseklik = Math.Max(
                    SERIT_ASGARI_YUKSEKLIK,
                    4 + satir1Yukseklik + (maddeVar ? 2 + maddeYukseklik : 0) + 4);
                int tuvalY = cerceveliY + seritYukseklik;

                var kare = new Bitmap(tuvalG, tuvalY, PixelFormat.Format32bppArgb);
                using (var g = Graphics.FromImage(kare))
                {
                    // Zemin: şerit rengi. Tuval içerikten genişse yanlarda
                    // paspartu gibi durur, hiç boş/şeffaf piksel kalmaz.
                    g.Clear(uyari ? SERIT_ZEMIN_UYARI : SERIT_ZEMIN);

                    int solKenar = (tuvalG - cerceveliG) / 2;

                    using (var firca = new SolidBrush(CERCEVE_RENK))
                        g.FillRectangle(firca, solKenar, 0, cerceveliG, cerceveliY);

                    g.CompositingMode = CompositingMode.SourceCopy;
                    g.DrawImage(icerik, solKenar + CERCEVE, CERCEVE, icerik.Width, icerik.Height);
                    g.CompositingMode = CompositingMode.SourceOver;

                    g.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;
                    using (var yaziFirca = new SolidBrush(SERIT_YAZI))
                    using (var maddeFirca = new SolidBrush(Color.FromArgb(255, 200, 200, 205)))
                    {
                        float y = cerceveliY + 4;

                        using (var bicim = new StringFormat(StringFormatFlags.NoWrap))
                        {
                            bicim.Alignment = StringAlignment.Center;
                            bicim.LineAlignment = StringAlignment.Near;
                            bicim.Trimming = StringTrimming.None;
                            g.DrawString(damgaSatiri, font1, yaziFirca,
                                new RectangleF(0, y, tuvalG, satir1Yukseklik), bicim);
                        }
                        y += satir1Yukseklik + 2;

                        if (maddeVar)
                        {
                            using (var bicim2 = new StringFormat())
                            {
                                bicim2.Alignment = StringAlignment.Center;
                                bicim2.LineAlignment = StringAlignment.Near;
                                bicim2.Trimming = StringTrimming.EllipsisCharacter;
                                g.DrawString(maddeSatiri, font2, maddeFirca,
                                    new RectangleF(SERIT_YAN_BOSLUK, y,
                                                   tuvalG - 2 * SERIT_YAN_BOSLUK, maddeYukseklik),
                                    bicim2);
                            }
                        }
                    }
                }
                return kare;
            }
            finally
            {
                if (font1 != null) font1.Dispose();
                if (font2 != null) font2.Dispose();
            }
        }

        /// <summary>
        /// Verilen genişliğe sığan en büyük yazı tipini seçer. Hiçbiri
        /// sığmazsa en küçüğüyle döner; çağıran tuvali genişletir.
        /// </summary>
        private static Font SeritFontu(Graphics g, string metin, int kullanilabilirGenislik,
                                       out SizeF olcu)
        {
            float[] boylar = { 10f, 9.5f, 9f, 8.5f, 8f, 7.5f };
            Font son = null;
            SizeF sonOlcu = SizeF.Empty;

            foreach (var boy in boylar)
            {
                var f = new Font("Segoe UI", boy, FontStyle.Regular, GraphicsUnit.Point);
                var o = g.MeasureString(metin, f);
                if (son != null) son.Dispose();
                son = f;
                sonOlcu = o;
                if (kullanilabilirGenislik > 0 && o.Width <= kullanilabilirGenislik) break;
            }

            olcu = sonOlcu;
            return son;
        }

        private static string BenzersizYol(string yol)
        {
            if (!File.Exists(yol)) return yol;

            var klasor = Path.GetDirectoryName(yol);
            var ad = Path.GetFileNameWithoutExtension(yol);
            var uzanti = Path.GetExtension(yol);
            for (int i = 2; i < 1000; i++)
            {
                var aday = Path.Combine(klasor, $"{ad}-{i}{uzanti}");
                if (!File.Exists(aday)) return aday;
            }
            return Path.Combine(klasor, $"{ad}-{Guid.NewGuid():N}{uzanti}");
        }

        /// <summary>
        /// Klasörü son <paramref name="sonKareSayisi"/> kareyle sınırlar.
        /// YALNIZ bu aracın ürettiği ada uyan dosyalara dokunur.
        /// </summary>
        /// <returns>silinen dosya sayısı</returns>
        public static int Buda(string klasor, int sonKareSayisi)
        {
            if (sonKareSayisi <= 0 || !Directory.Exists(klasor)) return 0;

            var kareler = new DirectoryInfo(klasor)
                .GetFiles("*.png")
                .Where(f => ReKareAdi.IsMatch(f.Name))
                .OrderByDescending(f => f.LastWriteTimeUtc)
                .ToList();

            int silinen = 0;
            for (int i = sonKareSayisi; i < kareler.Count; i++)
            {
                try { kareler[i].Delete(); silinen++; }
                catch { /* dosya kilitliyse sonraki turda silinir */ }
            }
            return silinen;
        }

        /// <summary>
        /// Bölge seçicinin döndürdüğü WPF dikdörtgenini gerçek EKRAN
        /// PİKSELİNE çevirir.
        ///
        /// ⚠️ RegionSelectorWindow karışık birim döndürüyor: X/Y için
        /// PointToScreen çağırdığı için onlar zaten CİHAZ pikseli, ama
        /// Width/Height WPF mantıksal birimi olarak kalıyor. %100 ölçekte
        /// ikisi eşit olduğu için fark görünmez; kullanıcı ölçeği %125'e
        /// alırsa kırpma küçük çıkar. Burada yalnız en/boy ölçeklenir.
        /// </summary>
        public static Rectangle WpfDikdortgenden(System.Windows.Rect r, double dpiX, double dpiY)
        {
            if (dpiX <= 0) dpiX = 1.0;
            if (dpiY <= 0) dpiY = 1.0;
            return new Rectangle(
                (int)Math.Round(r.X),
                (int)Math.Round(r.Y),
                Math.Max(1, (int)Math.Round(r.Width * dpiX)),
                Math.Max(1, (int)Math.Round(r.Height * dpiY)));
        }

        /// <summary>Varsayılan kare klasörü — OneDrive DIŞINDA.</summary>
        public static string VarsayilanKlasor()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TarihAtlasiKare");
        }
    }
}
