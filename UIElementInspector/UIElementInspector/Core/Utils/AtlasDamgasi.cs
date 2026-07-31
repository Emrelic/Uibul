using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
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

        /// <summary>
        /// Harita tuvalinin EKRAN pikselindeki dikdörtgeni: sol, üst, en, boy.
        /// Başlıkta "g:" ile gelir. Yoksa null.
        /// </summary>
        public int[] Tuval { get; private set; }

        /// <summary>
        /// Görünen alanın coğrafi sınırları: kuzey, batı, güney, doğu.
        /// Başlıkta "b:" ile gelir. Yoksa null.
        /// </summary>
        public double[] Sinir { get; private set; }

        /// <summary>Damganın okunduğu ham pencere başlığı (tanı için)</summary>
        public string HamBaslik { get; private set; }

        /// <summary>
        /// true ise başlıkta damga YOKTU; tarih/koordinat alanları gerçek
        /// değil, kare "damgasız" olarak üretildi. Şerit bu durumda kırmızı
        /// zeminle çizilir — sahte kesinlik izlenimi vermesin diye.
        /// </summary>
        public bool Eksik { get; private set; }

        private AtlasDamgasi() { }

        /// <summary>
        /// Başlıkta damga bulunamadığında kullanılan kare kimliği.
        ///
        /// Tarih UYDURULMAZ. Atlasın hangi güne bakıyor olduğu bilinmiyorsa
        /// yazılabilecek tek dürüst şey karenin NE ZAMAN alındığıdır; şerit
        /// bunu açıkça "tarih okunamadı" diyerek söyler. Alternatif —hiç kare
        /// almamak— kullanıcıyı aracı hiç kullanamaz hâle getiriyordu.
        /// </summary>
        public static AtlasDamgasi Damgasiz(string pencereBasligi, DateTime an)
        {
            var pencere = string.IsNullOrWhiteSpace(pencereBasligi)
                ? null
                : ReTarayiciKuyrugu.Replace(pencereBasligi, "").Trim();

            return new AtlasDamgasi
            {
                Tarih = an.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                Enlem = null,
                Boylam = null,
                Zoom = null,
                Madde = pencere,
                HamBaslik = pencereBasligi,
                Eksik = true,
                _damgasizAn = an
            };
        }

        private DateTime _damgasizAn;

        /// <summary>
        /// Görüntünün İÇİNE basılacak ve panoya metin olarak konacak satır.
        /// Dosya adı kaybolsa bile bilgi bu satırda taşınır — özelliğin
        /// varlık sebebi budur.
        /// </summary>
        public string Satir
        {
            get
            {
                if (Eksik)
                    return "TARİH/KOORDİNAT OKUNAMADI — kare " +
                           _damgasizAn.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture) +
                           " tarihinde alındı";
                return $"{Tarih} · {Enlem} {Boylam} · {Zoom} · Osmanlı Tarih Atlası";
            }
        }

        /// <summary>
        /// Şeridin birinci satırı, seçilen bölgenin koordinatıyla. Bölge
        /// hesaplanamadıysa (eski biçim başlık, ya da seçim haritanın dışında)
        /// merkez koordinatını yazan <see cref="Satir"/>'a düşer.
        /// </summary>
        public string SatirBolgeli(string bolgeYazisi)
        {
            if (Eksik || string.IsNullOrEmpty(bolgeYazisi)) return Satir;
            return $"{Tarih} · {bolgeYazisi} · {Zoom} · Osmanlı Tarih Atlası";
        }

        /// <summary>
        /// Şeridin ikinci satırı — açık kronoloji maddesi (damga varsa) ya da
        /// pencere başlığı (damga yoksa). Yoksa null.
        /// </summary>
        public string MaddeSatiri
        {
            get
            {
                if (string.IsNullOrWhiteSpace(Madde)) return null;
                return (Eksik ? "Pencere: " : "Madde: ") + Madde;
            }
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

        /// <summary>YYYY-AA-GG_enlem_boylam_zN.png (damga yoksa damgasiz_...)</summary>
        public string DosyaAdi
        {
            get
            {
                if (Eksik)
                    return "damgasiz_" +
                           _damgasizAn.ToString("yyyy-MM-dd_HHmmss", CultureInfo.InvariantCulture) +
                           ".png";
                return $"{Tarih}_{Enlem}_{Boylam}_{Zoom}.png";
            }
        }

        /// <summary>
        /// OKUNABİLİR dosya adı — kullanıcı isteği: dosyaya bakan biri açmadan
        /// hangi maddeyi, hangi tarihi ve nereyi gösterdiğini anlasın.
        ///
        ///   1281-01-01 · 39.34-40.77N 28.78-31.52E · Ertuğrul Gazi'nin ölümü.png
        ///
        /// <paramref name="bolgeYazisi"/> verilmezse merkez koordinatı yazılır.
        /// Windows'ta yasak karakterler ( \ / : * ? " &lt; &gt; | ) ayıklanır ve ad
        /// 150 karakterle sınırlanır — uzun madde başlıkları yol sınırını aşabilir.
        /// </summary>
        public string OkunakliDosyaAdi(string bolgeYazisi)
        {
            string govde;
            if (Eksik)
            {
                govde = "TARIHSIZ " +
                        _damgasizAn.ToString("yyyy-MM-dd HH-mm-ss", CultureInfo.InvariantCulture);
                if (!string.IsNullOrWhiteSpace(Madde)) govde += " · " + Madde;
            }
            else
            {
                var koordinat = string.IsNullOrWhiteSpace(bolgeYazisi)
                    ? Enlem + " " + Boylam
                    : bolgeYazisi.Replace(" · ", " ");
                govde = Tarih + " · " + koordinat;
                if (!string.IsNullOrWhiteSpace(Madde)) govde += " · " + Madde;
            }

            return Temizle(govde) + ".png";
        }

        private static string Temizle(string ad)
        {
            var sb = new StringBuilder(ad.Length);
            foreach (var c in ad)
            {
                // "–" (uzun tire) dosya adında geçerlidir ama bazı araçlar
                // bozuk gösteriyor; düz tireye indiriliyor.
                if (c == '–' || c == '—') { sb.Append('-'); continue; }
                if (Array.IndexOf(Path.GetInvalidFileNameChars(), c) >= 0) { sb.Append(' '); continue; }
                sb.Append(c);
            }

            var s = Regex.Replace(sb.ToString(), @"\s+", " ").Trim();
            s = s.TrimEnd('.', ' ');                       // Windows sondaki nokta/boşluğu atar
            if (s.Length > 150) s = s.Substring(0, 150).TrimEnd('.', ' ', '·');
            return s.Length == 0 ? "atlas-kare" : s;
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
                Tuval = TamsayiDizisi(baslik, "g:", 4),
                Sinir = SayiDizisi(baslik, "b:", 4),
                HamBaslik = baslik
            };
        }

        private static int[] TamsayiDizisi(string baslik, string onEk, int adet)
        {
            var d = SayiDizisi(baslik, onEk, adet);
            if (d == null) return null;
            var s = new int[adet];
            for (int i = 0; i < adet; i++) s[i] = (int)Math.Round(d[i]);
            return s;
        }

        // "g:8,131,1180,700" / "b:44.1000,20.3000,38.2000,32.9000" bölümlerini okur.
        private static double[] SayiDizisi(string baslik, string onEk, int adet)
        {
            foreach (var ham in baslik.Split('·'))
            {
                var b = ham.Trim();
                if (!b.StartsWith(onEk, StringComparison.OrdinalIgnoreCase)) continue;

                var parcalar = b.Substring(onEk.Length).Split(',');
                if (parcalar.Length != adet) return null;

                var d = new double[adet];
                for (int i = 0; i < adet; i++)
                {
                    d[i] = Sayi(parcalar[i].Trim());
                    if (double.IsNaN(d[i])) return null;
                }
                return d;
            }
            return null;
        }

        /// <summary>
        /// Ekranda seçilen dikdörtgenin coğrafi karşılığını yazar:
        /// "40.20–41.90N · 25.10–28.90E". Başlıkta g:/b: yoksa ya da seçim
        /// haritanın tamamen dışındaysa null.
        ///
        /// Doğruluk şartı: harita kuzeye bakmalı ve eğik olmamalı. app.js'te
        /// bearing/pitch hiç geçmiyor (ölçüldü), bu yüzden aşağıdaki ters
        /// Mercator dönüşümü yaklaşık değil TAM sonuç verir.
        /// </summary>
        public string BolgeYazisi(int secX, int secY, int secEn, int secBoy)
        {
            if (Tuval == null || Sinir == null) return null;

            double gx = Tuval[0], gy = Tuval[1], gw = Tuval[2], gh = Tuval[3];
            if (gw <= 0 || gh <= 0) return null;

            // Seçimi tuvalle kes — panel/araç çubuğu seçime girmiş olabilir.
            double x1 = Math.Max(secX, gx), x2 = Math.Min(secX + secEn, gx + gw);
            double y1 = Math.Max(secY, gy), y2 = Math.Min(secY + secBoy, gy + gh);
            if (x2 <= x1 || y2 <= y1) return null;   // seçim haritanın dışında

            double kuzey = Sinir[0], bati = Sinir[1], guney = Sinir[2], dogu = Sinir[3];

            // Boylam doğrusal
            double lonB = bati + (x1 - gx) / gw * (dogu - bati);
            double lonD = bati + (x2 - gx) / gw * (dogu - bati);

            // Enlem Mercator'da doğrusal
            double mK = Merkatore(kuzey), mG = Merkatore(guney);
            double lat1 = MerkatordenGeri(mK + (y1 - gy) / gh * (mG - mK));   // üst kenar
            double lat2 = MerkatordenGeri(mK + (y2 - gy) / gh * (mG - mK));   // alt kenar

            double enlemUst = Math.Max(lat1, lat2), enlemAlt = Math.Min(lat1, lat2);

            return Aralik(enlemAlt, enlemUst, "N", "S") + " · " +
                   Aralik(lonB, lonD, "E", "W");
        }

        private static double Merkatore(double enlem)
        {
            var r = Math.Max(-85.05, Math.Min(85.05, enlem)) * Math.PI / 180.0;
            return Math.Log(Math.Tan(Math.PI / 4.0 + r / 2.0));
        }

        private static double MerkatordenGeri(double m)
        {
            return (2.0 * Math.Atan(Math.Exp(m)) - Math.PI / 2.0) * 180.0 / Math.PI;
        }

        // "40.20–41.90N" — iki uç aynı yarımkürede ise tek harf, değilse ikisi de.
        private static string Aralik(double a, double b, string artiHarf, string eksiHarf)
        {
            string Yaz(double v) => Math.Abs(v).ToString("0.00", CultureInfo.InvariantCulture);

            if (a >= 0 && b >= 0) return Yaz(a) + "–" + Yaz(b) + artiHarf;
            if (a < 0 && b < 0) return Yaz(b) + "–" + Yaz(a) + eksiHarf;
            return Yaz(a) + (a < 0 ? eksiHarf : artiHarf) + "–" +
                   Yaz(b) + (b < 0 ? eksiHarf : artiHarf);
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
                // Makine alanları: g: tuval dikdörtgeni, b: görünen alan sınırları
                if (b.StartsWith("g:", StringComparison.OrdinalIgnoreCase)) continue;
                if (b.StartsWith("b:", StringComparison.OrdinalIgnoreCase)) continue;
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
            string yoksay;
            return Oku(out tani, out yoksay);
        }

        /// <summary>
        /// <paramref name="bulunanBaslik"/>: damga çıkmasa bile eldeki en
        /// anlamlı pencere başlığı (varsa atlas penceresininki, yoksa ön
        /// plandakininki). Damgasız kare bunu şeride yazar.
        /// </summary>
        public static AtlasDamgasi Oku(out string tani, out string bulunanBaslik)
        {
            var onPlan = BaslikAl(GetForegroundWindow());
            bulunanBaslik = onPlan;

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
            {
                bulunanBaslik = bulunanAtlas;
                tani = $"atlas penceresi bulundu ama başlıkta damga YOK: [{bulunanAtlas}]";
            }
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
