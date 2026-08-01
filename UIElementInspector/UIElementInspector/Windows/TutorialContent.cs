using System.Collections.Generic;

namespace UIElementInspector.Windows
{
    public enum BlokTuru
    {
        Baslik,     // ara başlık
        Paragraf,   // düz metin
        Madde,      // • listesi
        Numarali,   // 1. 2. 3. adım listesi
        Tus,        // klavye tuşu kartı
        Ipucu,      // yeşil kutu
        Uyari,      // turuncu/kırmızı kutu
        Bilgi,      // mavi kutu
        Ornek       // gri "senaryo" kutusu
    }

    public sealed class Blok
    {
        public BlokTuru Tur { get; init; }
        public string Metin { get; init; } = "";
        public string Etiket { get; init; } = "";
        public IReadOnlyList<string> Ogeler { get; init; } = System.Array.Empty<string>();

        public static Blok H(string m) => new() { Tur = BlokTuru.Baslik, Metin = m };
        public static Blok P(string m) => new() { Tur = BlokTuru.Paragraf, Metin = m };
        public static Blok M(params string[] o) => new() { Tur = BlokTuru.Madde, Ogeler = o };
        public static Blok N(params string[] o) => new() { Tur = BlokTuru.Numarali, Ogeler = o };
        public static Blok T(string tus, string m) => new() { Tur = BlokTuru.Tus, Etiket = tus, Metin = m };
        public static Blok Ipucu(string m) => new() { Tur = BlokTuru.Ipucu, Metin = m };
        public static Blok Uyari(string m) => new() { Tur = BlokTuru.Uyari, Metin = m };
        public static Blok Bilgi(string m) => new() { Tur = BlokTuru.Bilgi, Metin = m };
        public static Blok Ornek(string baslik, string m) => new() { Tur = BlokTuru.Ornek, Etiket = baslik, Metin = m };
    }

    public sealed class Adim
    {
        public string Bolum { get; init; } = "";
        public string Baslik { get; init; } = "";
        public string Ozet { get; init; } = "";
        public IReadOnlyList<Blok> Bloklar { get; init; } = System.Array.Empty<Blok>();
    }

    /// <summary>
    /// Öğreticinin tüm içeriği. Kod değil, metindir — yeni bir özellik
    /// eklendiğinde buraya bir <see cref="Adim"/> eklemek yeterlidir.
    /// </summary>
    public static class TutorialContent
    {
        public static IReadOnlyList<Adim> Adimlar { get; } = new List<Adim>
        {
            // ══════════════ BÖLÜM 1 — TANIŞMA ══════════════
            new Adim
            {
                Bolum = "1 · TANIŞMA",
                Baslik = "UIBUL nedir?",
                Ozet = "Ekranda gördüğünüz her şeyi okuyabilen, yakalayabilen ve belgeleyebilen bir araç.",
                Bloklar = new[]
                {
                    Blok.P("UIBUL, Windows'ta çalışan her uygulamanın ve her web sayfasının arayüzünü " +
                           "\"içeriden\" okuyabilen bir inceleme aracıdır. Ekranda gördüğünüz bir düğmenin " +
                           "sadece resmini değil; adını, tipini, kimliğini, konumunu, HTML'ini ve " +
                           "erişilebilirlik bilgilerini de alır."),
                    Blok.P("Ama aynı zamanda çok daha basit bir şey: hızlı bir ekran görüntüsü ve " +
                           "arşivleme aracıdır. İki yüzü olan bir programdır."),
                    Blok.H("İki farklı kullanıcı, aynı program"),
                    Blok.M(
                        "YAZILIMCI: test otomasyonu için selector toplar, hata raporu hazırlar, erişilebilirlik denetler, eski uygulamaları çözer.",
                        "HERKES: F9 ile saniyede bölge görüntüsü alır, dosya yollarını yapıştırır, ekran kanıtlarını tarih damgalı arşivler."),
                    Blok.Bilgi("Bu öğretici ikisini de anlatır. Yazılımcı değilseniz 4. bölüm " +
                               "(Günlük Kullanım) sizin için en değerli kısım — isterseniz soldaki " +
                               "listeden doğrudan oraya atlayabilirsiniz."),
                    Blok.Ipucu("Öğreticiyi istediğiniz zaman kapatabilirsiniz. Tekrar açmak için: " +
                               "üst şeritteki mor 🎓 ÖĞRETİCİ düğmesi ya da Help ▸ Öğretici.")
                }
            },

            new Adim
            {
                Bolum = "1 · TANIŞMA",
                Baslik = "Program penceresinin anatomisi",
                Ozet = "Neresi ne işe yarıyor — pencereyi beş parçaya bölerek bakalım.",
                Bloklar = new[]
                {
                    Blok.H("① En üst: menü çubuğu"),
                    Blok.P("File / Edit / View / Tools / Help. Klasik menü. Export, ayarlar, " +
                           "öğretici ve güncelleme kontrolü buradan da açılır."),

                    Blok.H("② Üst şerit: büyük renkli düğmeler"),
                    Blok.P("En sık kullanılan işlemler: BAŞLAT (F1), GÖRÜNÜR BAŞLAT (F3), DURDUR (F2), " +
                           "YENİLE (F5), KAYDET, ÖĞRETİCİ, KILAVUZ ve en sağda 📌 HER ZAMAN ÜSTTE. " +
                           "Her düğmenin üzerine gelirseniz ne yaptığını anlatan bir balon çıkar."),

                    Blok.H("③ Sol panel: toplanan elementler"),
                    Blok.P("Yakaladığınız her element buraya bir satır olarak düşer. Ağaç görünümünde " +
                           "parent-child ilişkisini de gösterir. Arama kutusuyla içinde arayabilirsiniz."),

                    Blok.H("④ Sağ panel: sekmeler"),
                    Blok.M(
                        "RAW PROPERTIES — elementin ham özellik dökümü",
                        "ALL PROPERTIES — 10 kategoriye ayrılmış tam liste (Basic, UIA, CDP/Web, Win32, Position, Selectors, Hierarchy, Table/Grid, State)",
                        "SOURCE CODE — web elementinin HTML kaynağı",
                        "SCREENSHOT — o elementin/pencerenin görüntüsü",
                        "ARCHIVE — geçmiş yakalamalarınızın arşivi"),

                    Blok.H("⑤ Alt: kısayol çubuğu ve konsol"),
                    Blok.P("Pencerenin altında F1'den F11'e kadar tüm kısayolların düğmeleri sıralanır. " +
                           "Tuşu hatırlamıyorsanız düğmeye tıklayabilirsiniz. Altındaki konsol ise " +
                           "programın ne yaptığını satır satır yazar — bir şey çalışmadığında ilk " +
                           "bakılacak yer burasıdır."),

                    Blok.Ipucu("Sağ üstteki 📌 HER ZAMAN ÜSTTE düğmesi, pencereyi diğer tüm " +
                               "pencerelerin üstünde tutar. İnceleme yaparken çok işe yarar.")
                }
            },

            // ══════════════ BÖLÜM 2 — İNCELEME ══════════════
            new Adim
            {
                Bolum = "2 · ELEMENT İNCELEME",
                Baslik = "F1 — İncelemeyi başlat",
                Ozet = "Programın kalbi. Basınca UIBUL kaybolur ve ekranda gezindiğiniz her şeyi okumaya başlar.",
                Bloklar = new[]
                {
                    Blok.T("F1", "İncelemeyi başlat — ana pencere gizlenir"),
                    Blok.N(
                        "İncelemek istediğiniz uygulamayı veya web sayfasını açın.",
                        "F1'e basın. UIBUL penceresi kaybolur, sağ üstte küçük bir kontrol paneli belirir.",
                        "Mouse'u incelemek istediğiniz düğmenin/kutunun üzerine getirin.",
                        "Element anında algılanır; küçük panelde sayaç artar.",
                        "F2 ile durdurun — pencere geri gelir, topladığınız her şey içindedir."),
                    Blok.Bilgi("Pencerenin gizlenmesi kasıtlıdır: incelemek istediğiniz şeyin önünü " +
                               "kapatmasın diye. Ekran görüntüsü alırken de UIBUL görüntüye girmez."),
                    Blok.Uyari("F1 ve F2 GLOBAL kısayollardır — UIBUL arka planda, hatta gizliyken bile " +
                               "çalışırlar. Başka bir uygulamadayken F2'ye basmanız yeter."),
                    Blok.Ipucu("Bir şeyi \"tam olarak neresi\" diye merak ediyorsanız F1 ile gezinip " +
                               "F2'ye basın; sol panelde gezdiğiniz her elementin listesi durur.")
                }
            },

            new Adim
            {
                Bolum = "2 · ELEMENT İNCELEME",
                Baslik = "F3 — Görünür başlat",
                Ozet = "Aynı inceleme, ama pencere gizlenmez. Sonucu anında görmek isteyenler için.",
                Bloklar = new[]
                {
                    Blok.T("F3", "İncelemeyi başlat — pencere görünür kalır"),
                    Blok.P("F1 pencereyi gizler, F3 gizlemez. Fark bu kadar basit."),
                    Blok.H("F3 ne zaman daha iyi?"),
                    Blok.M(
                        "İki monitörünüz varsa: hedef uygulama birinde, UIBUL diğerinde açık kalır.",
                        "Element özelliklerini anında okumak istiyorsanız — panel gözünüzün önünde durur.",
                        "Küçük bir şeye hızlı bakacaksanız, F2'ye basıp pencereyi geri çağırmakla uğraşmazsınız."),
                    Blok.H("F1 ne zaman daha iyi?"),
                    Blok.M(
                        "Tek monitörde çalışıyorsanız.",
                        "İnceleyeceğiniz alan ekranın büyük kısmını kaplıyorsa.",
                        "İnceleme sırasında ekran görüntüsü de alacaksanız.")
                }
            },

            new Adim
            {
                Bolum = "2 · ELEMENT İNCELEME",
                Baslik = "F4 — Deklanşör: kaybolan menüleri yakalama",
                Ozet = "Açılır menüler mouse'u çekince kapanır. F4 bu sorunu çözer.",
                Bloklar = new[]
                {
                    Blok.T("F4", "Basılı tut → hedefe git → bırak = yakala"),
                    Blok.P("Klasik problem: bir sağ tık menüsünü veya dropdown listesini incelemek " +
                           "istersiniz, ama UIBUL'a dönmek için mouse'u oynattığınızda menü kapanır. " +
                           "F4 tam olarak bunun içindir — fotoğraf makinesinin deklanşörü gibi çalışır."),
                    Blok.N(
                        "Menüyü/dropdown'ı açın; açık kalsın.",
                        "F4 tuşuna BASILI TUTUN (bırakmayın).",
                        "Basılı tutarken mouse'u menüdeki öğenin üzerine götürün.",
                        "F4'ü BIRAKIN — tam o andaki element yakalanır."),
                    Blok.Ipucu("Sağ tık menüleri, otomatik tamamlama listeleri, tooltip'ler, tarih " +
                               "seçiciler ve hover ile açılan alt menüler için tek pratik yol budur."),
                    Blok.Ornek("Gerçek senaryo",
                        "Bir web sitesindeki dropdown'ın seçeneklerinin gerçek value'larını öğrenmeniz " +
                        "gerekiyor. Dropdown'ı açın, F4'ü basılı tutun, seçeneğin üzerine gelin, " +
                        "bırakın. Value, index ve seçici bilgisi elinizde.")
                }
            },

            new Adim
            {
                Bolum = "2 · ELEMENT İNCELEME",
                Baslik = "F2 ve F5 — Durdur ve yenile",
                Ozet = "İki küçük ama sürekli kullanılan tuş.",
                Bloklar = new[]
                {
                    Blok.T("F2", "İncelemeyi durdur, pencereyi geri getir"),
                    Blok.P("Toplanan veriler silinmez; sol panelde durur ve incelemeye hazırdır. " +
                           "F2 global çalışır: hangi uygulamada olursanız olun basabilirsiniz."),
                    Blok.T("F5", "Seçili elementi yeniden analiz et"),
                    Blok.P("Sayfa JavaScript ile değiştiyse, bir alan dolduysa ya da durum " +
                           "(enabled/checked) değiştiyse F5 elementi baştan okur. Dinamik sayfalarda " +
                           "\"bu bilgi eski mi?\" diye şüphelendiğinizde ilk yapacağınız şey."),
                    Blok.Uyari("F5'in yenilediği şey SEÇİLİ elementtir — listenin tamamı değil. " +
                               "Tüm listeyi tazelemek için incelemeyi yeniden çalıştırmanız gerekir.")
                }
            },

            new Adim
            {
                Bolum = "2 · ELEMENT İNCELEME",
                Baslik = "Sağdaki sekmeler ne gösterir?",
                Ozet = "Aynı elementin beş farklı yüzü.",
                Bloklar = new[]
                {
                    Blok.H("RAW PROPERTIES"),
                    Blok.P("Ham döküm. Hızlı bakış için; kopyalayıp bir yere yapıştırmaya en uygun biçim."),

                    Blok.H("ALL PROPERTIES"),
                    Blok.P("İşin ciddi kısmı. 10 kategoriye ayrılmış tam liste:"),
                    Blok.M(
                        "1. BASIC — ad, tip, değer, sınıf, kimlik",
                        "2. UIA — AutomationId, ControlType, desteklenen pattern'ler",
                        "3. CDP/Web — tarayıcıdan gelen DOM bilgisi",
                        "5. Win32 — pencere sınıfı, handle, stil bayrakları",
                        "6. Position — X, Y, genişlik, yükseklik",
                        "7. Selectors — XPath ve CSS seçicileri (otomasyon için altın değerinde)",
                        "8. Hierarchy — üst/alt element ilişkisi",
                        "9. Table/Grid — tablo ise satır/sütun bilgisi",
                        "10. State — enabled, visible, focused, checked"),

                    Blok.H("SOURCE CODE"),
                    Blok.P("Web elementinin HTML kaynağı. Masaüstü uygulamalarında boş kalabilir — normaldir."),

                    Blok.H("SCREENSHOT"),
                    Blok.P("Elementin görüntüsü. F7 ile yakaladıysanız üç ayrı görüntü olur: " +
                           "tam ekran, pencere ve elementin kendisi."),

                    Blok.H("ARCHIVE"),
                    Blok.P("Geçmiş yakalamalarınız. Arşive attığınız her şey burada kalır, " +
                           "yeniden adlandırabilir veya silebilirsiniz.")
                }
            },

            // ══════════════ BÖLÜM 3 — YAKALAMA VE KAYIT ══════════════
            new Adim
            {
                Bolum = "3 · YAKALAMA VE KAYIT",
                Baslik = "F7 — Tam yakalama (en güçlü tuş)",
                Ozet = "Tek tuşta: 5 teknolojiyle analiz + DOM ağacı + kaynak kod + 3 ekran görüntüsü.",
                Bloklar = new[]
                {
                    Blok.T("F7", "Tam yakalama → masaüstü + arşiv"),
                    Blok.P("F7, elinizdeki en kapsamlı komuttur. Bastığınızda şunların hepsi bir " +
                           "klasöre paketlenir:"),
                    Blok.M(
                        "Beş algılama teknolojisinin (UIA, CDP, MSHTML, Win32, Playwright) ayrı ayrı sonucu",
                        "Sayfa/pencere yapısı — DOM ağacı",
                        "Kaynak kodlar",
                        "Üç ekran görüntüsü: tam ekran, aktif pencere, elementin kendisi"),
                    Blok.P("Çıktı hem masaüstüne hem arşive yazılır. Klasör adı okunabilir: " +
                           "tarih, saat ve içerik bilgisi içerir."),
                    Blok.Uyari("F7 kapsamlı olduğu için birkaç saniye sürebilir. Ekranda dairesel bir " +
                               "ilerleme göstergesi çıkar — bitmesini bekleyin."),
                    Blok.Ipucu("Bir hatayı birine anlatmanız gerekiyorsa F7 en iyi yatırımdır: " +
                               "klasörü zipleyip gönderin, karşı taraf her şeyi görür.")
                }
            },

            new Adim
            {
                Bolum = "3 · YAKALAMA VE KAYIT",
                Baslik = "F8, F6 ve Ctrl+S — daha hafif kayıtlar",
                Ozet = "Her zaman tam paket gerekmez. Üç kademe daha var.",
                Bloklar = new[]
                {
                    Blok.T("F8", "F7'nin aynısı, ama SADECE arşive"),
                    Blok.P("Masaüstünüzü kirletmeden çalışmak istiyorsanız F8 kullanın. İçerik " +
                           "birebir aynıdır, yalnızca hedef farklıdır. Uzun bir inceleme " +
                           "seansında onlarca yakalama yapıyorsanız F8 tercih edilir."),

                    Blok.T("F6", "TXT rapor → masaüstü + arşiv"),
                    Blok.P("Ekran görüntüsü ve kaynak kod olmadan, sadece okunabilir bir metin raporu. " +
                           "Hızlıdır, küçüktür, e-postaya yapıştırmaya uygundur."),

                    Blok.T("Ctrl+S", "Hızlı kaydet"),
                    Blok.P("En hızlı yol. Mevcut element bilgisini anında masaüstüne TXT olarak atar, " +
                           "hiçbir soru sormaz."),

                    Blok.H("Hangisini ne zaman?"),
                    Blok.M(
                        "Sorunu birine göstereceksem → F7",
                        "Kendi arşivim için topluyorsam → F8",
                        "Sadece bilgi lazımsa → F6",
                        "Acelem varsa → Ctrl+S")
                }
            },

            new Adim
            {
                Bolum = "3 · YAKALAMA VE KAYIT",
                Baslik = "Export biçimleri — hangisi ne için?",
                Ozet = "CSV, JSON, XML, HTML, TXT. Seçim, veriyi nereye götüreceğinize bağlı.",
                Bloklar = new[]
                {
                    Blok.M(
                        "CSV — Excel'de açmak, filtrelemek, sıralamak için. Toplu element listelerinde en pratiği.",
                        "JSON — kod içinde işlemek için. Test scriptlerinize doğrudan besleyebilirsiniz.",
                        "XML — kurumsal sistemlere, şema doğrulaması gereken yerlere.",
                        "HTML — birine göstermek için. Tarayıcıda açılır, filtrelenebilir tablo olur.",
                        "TXT — okumak, e-postaya yapıştırmak, log tutmak için."),
                    Blok.P("Menüden: File ▸ Export. Ya da Ctrl+S ile varsayılan biçimde hızlı kayıt."),
                    Blok.Ipucu("Varsayılan çıktı klasörünü Tools ▸ Settings'ten değiştirebilirsiniz. " +
                               "Tarihe göre alt klasör açma seçeneği de oradadır."),
                    Blok.Ornek("Gerçek senaryo",
                        "Bir formdaki 40 alanın tamamının adını ve id'sini bir tabloya dökmeniz " +
                        "gerekiyor: F1 ile form üzerinde gezin, F2, sonra CSV export. Excel'de açın, " +
                        "işiniz bitti.")
                }
            },

            // ══════════════ BÖLÜM 4 — GÜNLÜK KULLANIM ══════════════
            new Adim
            {
                Bolum = "4 · GÜNLÜK KULLANIM",
                Baslik = "F9 — Bölge ekran görüntüsü",
                Ozet = "Yazılımla hiç ilgisi olmayan, ama en çok kullanacağınız tuş.",
                Bloklar = new[]
                {
                    Blok.T("F9", "Mouse ile bölge seç → PNG"),
                    Blok.P("F9, UIBUL'un \"herkesin aracı\" tarafıdır. Windows'un kendi Ekran Alıntısı " +
                           "aracına benzer ama üç farkı vardır:"),
                    Blok.M(
                        "Dosya OTOMATİK kaydedilir — \"kaydet\" penceresiyle uğraşmazsınız.",
                        "Panoya hem RESİM hem DOSYA YOLU kopyalanır — Word'e de yapıştırabilirsiniz, sohbete de.",
                        "Global çalışır — hangi programda olursanız olun F9 çalışır."),
                    Blok.N(
                        "F9'a basın. Ekran hafifçe kararır.",
                        "Mouse ile istediğiniz alanı sürükleyerek seçin.",
                        "Bırakın. Görüntü masaüstünüze PNG olarak kaydedilir ve panoya kopyalanır.",
                        "İstediğiniz yere Ctrl+V ile yapıştırın."),
                    Blok.Ipucu("Vazgeçmek için ESC. Seçim yaparken koordinatlar canlı gösterilir."),
                    Blok.Uyari("F9 GLOBAL kısayoldur; UIBUL açık olduğu sürece başka programlarda " +
                               "F9'a atanmış işlevler çalışmayabilir. Rahatsız ederse UIBUL'u kapatmanız yeter.")
                }
            },

            new Adim
            {
                Bolum = "4 · GÜNLÜK KULLANIM",
                Baslik = "Günlük hayatta F9 ile neler yapılır?",
                Ozet = "Somut örnekler — hiçbiri yazılım işi değil.",
                Bloklar = new[]
                {
                    Blok.Ornek("Destek talebi",
                        "Bankanın sitesinde bir hata aldınız. F9 ile hata mesajını seçin, " +
                        "destek sohbetine doğrudan Ctrl+V. Ne olduğunu anlatmaya çalışmaktan kurtulursunuz."),
                    Blok.Ornek("Fatura ve dekont arşivi",
                        "Online ödeme yaptınız. Dekont ekranını F9 ile alın. Dosya adında tarih " +
                        "olduğu için sonradan aramak kolaydır."),
                    Blok.Ornek("Alışverişte fiyat kanıtı",
                        "Bir ürün indirimdeyken F9 ile fiyatı kaydedin. İndirim iddiası tutmazsa " +
                        "elinizde tarihli kanıt olur."),
                    Blok.Ornek("Ders / toplantı notu",
                        "Ekrandaki bir grafiği ya da slaytı F9 ile alıp not defterinize yapıştırın. " +
                        "Fotoğraf çekmekten çok daha temiz olur."),
                    Blok.Ornek("Form doldurma yardımı",
                        "Uzun bir başvuru formunu doldururken, doldurduğunuz bölümleri F9 ile " +
                        "kaydedin. Sayfa çökerse ne yazdığınızı hatırlarsınız."),
                    Blok.Ornek("Rezervasyon / bilet",
                        "Uçuş, otel ya da randevu onay ekranını F9 ile alın. " +
                        "E-posta gelmezse elinizde kayıt olur."),
                    Blok.Ipucu("F9 ile alınan görüntüler masaüstüne düşer. Karışmasını istemiyorsanız " +
                               "Tools ▸ Settings'ten çıktı klasörünü değiştirin.")
                }
            },

            new Adim
            {
                Bolum = "4 · GÜNLÜK KULLANIM",
                Baslik = "F10 — Son yakalama yolunu yapıştır",
                Ozet = "\"Dosyayı nereye kaydetti?\" sorusunun cevabı, tek tuşla.",
                Bloklar = new[]
                {
                    Blok.T("F10", "Son yakalama klasörünün yolunu panoya kopyala ve yapıştır"),
                    Blok.P("F7/F8/F9 ile bir şey yakaladınız. Şimdi o klasörü birine göndermek ya da " +
                           "bir programa açtırmak istiyorsunuz. F10, yolu panoya koyar ve aktif " +
                           "pencereye yapıştırmayı dener."),
                    Blok.N(
                        "Önce F7, F8 veya F9 ile bir yakalama yapın.",
                        "Yolu yazmak istediğiniz yere gidin (sohbet kutusu, Explorer adres çubuğu, terminal).",
                        "F10'a basın."),
                    Blok.Uyari("UIBUL penceresi öndeyken F10 otomatik yapıştırmaz — yol yalnızca " +
                               "panoya konur. Hedef pencereye geçip Ctrl+V yapmanız gerekir. " +
                               "Bu kasıtlıdır: yanlış yere yazmasın diye."),
                    Blok.Ornek("Gerçek senaryo",
                        "Ekran görüntüsünü bir arkadaşınıza WhatsApp Web'den göndereceksiniz. " +
                        "F9 ile alın, WhatsApp'a geçin, F10'a basın — dosya yolu yazılır, " +
                        "oradan dosyayı seçmek saniyeler alır.")
                }
            },

            new Adim
            {
                Bolum = "4 · GÜNLÜK KULLANIM",
                Baslik = "F11 — Kimlik şeritli kare (Atlas)",
                Ozet = "Görüntünün içine tarih ve konum basan özel bir yakalama.",
                Bloklar = new[]
                {
                    Blok.T("F11", "Bölge seç → kırmızı çerçeve + kimlik şeridi"),
                    Blok.P("F11 başlangıçta Osmanlı Tarih Atlası projesindeki bir kusuru bildirmek için " +
                           "yapıldı, ama mantığı geneldir: aldığınız karenin ÜZERİNE, o karenin nereden " +
                           "ve ne zaman alındığını yazar."),
                    Blok.M(
                        "Kırpılan alanın etrafına 3 piksel kırmızı çerçeve çizilir.",
                        "Altına ince bir şerit basılır: tarih · koordinat · zoom · ilgili kayıt.",
                        "Bu bilgi görüntünün İÇİNDEDİR — dosya adı kaybolsa bile bilgi kareyle birlikte gider.",
                        "Kare hem panoya kopyalanır hem diske kaydedilir."),
                    Blok.Bilgi("Bilgi, tarayıcı penceresinin başlığından okunur. Başlıkta uygun damga " +
                               "yoksa kare yine alınır; şerit kırmızı zeminle 'TARİH/KOORDİNAT OKUNAMADI' " +
                               "der ve karenin alındığı saati yazar. Tarih asla uydurulmaz."),
                    Blok.Uyari("F11 global kısayol olarak kaydedildiği için, UIBUL açıkken Chrome'un " +
                               "F11 tam ekran kısayolu ÇALIŞMAZ. Rahatsız ederse ayar dosyasında " +
                               "AtlasKisayolu değerini \"Ctrl+F11\" yapın."),
                    Blok.Ipucu("Genel amaçlı kullanım: bir web sayfasından tarihli kanıt alırken " +
                               "F11 tercih edin — kimin, ne zaman, nereden aldığı görüntünün içinde durur.")
                }
            },

            // ══════════════ BÖLÜM 5 — YAZILIMCI ══════════════
            new Adim
            {
                Bolum = "5 · YAZILIMCI İÇİN",
                Baslik = "Test otomasyonu: selector avı",
                Ozet = "UIBUL'un bir yazılımcıya en somut faydası.",
                Bloklar = new[]
                {
                    Blok.P("Selenium, Playwright, Appium veya WinAppDriver ile test yazarken en çok " +
                           "vakit alan iş, elementleri güvenilir biçimde bulmaktır. Tarayıcının " +
                           "DevTools'u web için iş görür; masaüstü uygulamalarında ise elinizde " +
                           "genellikle hiçbir şey yoktur."),
                    Blok.H("UIBUL ne kazandırır?"),
                    Blok.M(
                        "Web ve masaüstü için TEK araç — iki ayrı alete gerek kalmaz.",
                        "ALL PROPERTIES ▸ 7. Selectors sekmesinde hazır XPath ve CSS seçicileri.",
                        "AutomationId, ControlType ve desteklenen pattern'ler — WinAppDriver için gereken tam da budur.",
                        "Toplu toplama: bir ekrandaki tüm elementleri tek seferde alıp CSV/JSON'a dökebilirsiniz."),
                    Blok.Ornek("İş akışı",
                        "1. Test edilecek ekranı açın.\n" +
                        "2. F1 ile inceleme başlatın, etkileşilecek elementler üzerinde gezin.\n" +
                        "3. F2 ile durdurun.\n" +
                        "4. JSON export alın.\n" +
                        "5. JSON'daki AutomationId / XPath değerlerini Page Object sınıfınıza dökün."),
                    Blok.Ipucu("Dinamik id üreten uygulamalarda F5 ile elementi birkaç kez yenileyip " +
                               "hangi alanın sabit kaldığına bakın — kararlı seçiciyi böyle bulursunuz.")
                }
            },

            new Adim
            {
                Bolum = "5 · YAZILIMCI İÇİN",
                Baslik = "Erişilebilirlik denetimi ve hata raporu",
                Ozet = "İki farklı iş, aynı araçla.",
                Bloklar = new[]
                {
                    Blok.H("Erişilebilirlik (a11y)"),
                    Blok.P("Ekran okuyucuların gördüğü şey, UIBUL'un okuduğu şeyin aynısıdır — ikisi de " +
                           "UI Automation / ARIA katmanından besleniyor. Bir düğmenin adı boşsa, " +
                           "rolü yanlışsa ya da etiketi yoksa UIBUL bunu size gösterir."),
                    Blok.M(
                        "Collection Profile'ı Full yapın.",
                        "F1 ile tüm etkileşimli elementler üzerinden geçin.",
                        "HTML export alın — filtrelenebilir bir denetim raporu olur.",
                        "Adı boş, rolü generic olan elementleri işaretleyin."),

                    Blok.H("Hata raporu (bug report)"),
                    Blok.P("İyi bir hata raporu üç şey ister: ne göründüğü, hangi elementte olduğu, " +
                           "hangi ortamda olduğu. F7 üçünü birden tek klasöre koyar."),
                    Blok.Ornek("Pratik",
                        "Hatayı ekranda yakaladığınız anda F7'ye basın. Oluşan klasörü zipleyin ve " +
                        "issue'ya ekleyin. Ekran görüntüsü, DOM, element özellikleri ve pencere " +
                        "bilgisi paketin içindedir — karşı taraf \"tekrar edemiyorum\" diyemez."),
                    Blok.Ipucu("Ekibinizde UIBUL herkeste varsa, hata raporu formatınızı " +
                               "\"F7 klasörünü ekle\" diye standartlaştırabilirsiniz.")
                }
            },

            new Adim
            {
                Bolum = "5 · YAZILIMCI İÇİN",
                Baslik = "Eski sistemler ve yapay zekâya bağlam verme",
                Ozet = "Belgesi olmayan uygulamaları çözmek ve LLM'lere ekranı anlatmak.",
                Bloklar = new[]
                {
                    Blok.H("Belgesiz / eski (legacy) uygulamalar"),
                    Blok.P("Kaynak kodu olmayan, kimsenin hatırlamadığı bir iç uygulamayı otomatikleştirmeniz " +
                           "gerekiyor. UIBUL, pencere sınıflarını, kontrol hiyerarşisini ve handle'ları " +
                           "çıkarır — uygulamanın iç yapısını dışarıdan haritalandırırsınız."),
                    Blok.M(
                        "Full Window modu ile tüm pencerenin element ağacını çıkarın.",
                        "Win32 kategorisinden pencere sınıflarını okuyun.",
                        "Hierarchy ile hangi kontrolün nerede olduğunu görün.",
                        "XML/JSON export ile bu haritayı takıma bırakın."),

                    Blok.H("Yapay zekâya bağlam vermek"),
                    Blok.P("Bir LLM'e \"şu ekranda şu düğmeye tıklayan kodu yaz\" derken en büyük sorun, " +
                           "modelin ekranı görmemesidir. F7 çıktısı tam olarak bu boşluğu doldurur: " +
                           "görüntü + element özellikleri + seçiciler birlikte gider."),
                    Blok.Ipucu("F11'in kimlik şeridi mantığı burada da işe yarar: görüntünün içine " +
                               "basılan bilgi, dosya adı kaybolsa bile modele hangi ekranın " +
                               "konuşulduğunu söyler.")
                }
            },

            // ══════════════ BÖLÜM 6 — AYARLAR VE BAKIM ══════════════
            new Adim
            {
                Bolum = "6 · AYARLAR VE BAKIM",
                Baslik = "Ayarlar — nerede ne değişir?",
                Ozet = "Tools ▸ Settings ve ayar dosyası.",
                Bloklar = new[]
                {
                    Blok.H("Tools ▸ Settings penceresinden"),
                    Blok.M(
                        "Çıktı klasörü — masaüstü dolmasın istiyorsanız ilk değiştireceğiniz ayar",
                        "Tarihe göre alt klasör açma",
                        "Collection Profile — Quick / Standard / Full",
                        "Ekran görüntüsü biçimi",
                        "Bildirimler ve günlük (log) ayarları"),

                    Blok.H("Ayar dosyası (ileri seviye)"),
                    Blok.P("Tüm ayarlar şu dosyada JSON olarak durur:"),
                    Blok.P("%AppData%\\UIElementInspector\\settings.json"),
                    Blok.M(
                        "AtlasKisayolu — F11 yerine başka bir tuş",
                        "AtlasKlasoru — atlas karelerinin kaydedileceği yer",
                        "AtlasEnUzunKenar — görüntü küçültme sınırı (varsayılan 1200 px)",
                        "OtomatikGuncellemeKontrolu — açılışta güncelleme baksın mı",
                        "GuncellemeDeposu — güncellemelerin çekileceği GitHub deposu"),
                    Blok.Uyari("Dosyayı elle düzenleyecekseniz UIBUL kapalıyken yapın; " +
                               "yoksa program kapanırken üzerine yazabilir.")
                }
            },

            new Adim
            {
                Bolum = "6 · AYARLAR VE BAKIM",
                Baslik = "Güncellemeler nasıl gelir?",
                Ozet = "Siz bir şey yapmadan yeni sürümler bulunur ve tek tıkla kurulur.",
                Bloklar = new[]
                {
                    Blok.P("UIBUL, açılışta (günde en fazla bir kez) GitHub'daki yayın sayfasına bakar. " +
                           "Yeni bir sürüm varsa bir pencere açılır; yoksa hiçbir şey olmaz."),
                    Blok.N(
                        "Güncelleme penceresi açılır ve yeniliklerin listesini gösterir.",
                        "\"Şimdi güncelle\" derseniz kurulum dosyası indirilir (ilerleme çubuğuyla).",
                        "İndirme bitince UIBUL kapanır, güncelleme kurulur ve program yeniden açılır.",
                        "Ayarlarınız ve arşiviniz korunur."),
                    Blok.H("Kontrolü siz yapmak isterseniz"),
                    Blok.P("Help ▸ Güncellemeleri kontrol et. Bu yol her zaman sonuç gösterir — " +
                           "güncel olsanız bile size bunu söyler."),
                    Blok.Bilgi("İnternet yoksa ya da GitHub'a ulaşılamıyorsa açılışta hata " +
                               "penceresi ÇIKMAZ; durum sadece alttaki konsola yazılır. " +
                               "Uygulama internetsiz de sorunsuz çalışır."),
                    Blok.Ipucu("Bir sürümü şimdilik istemiyorsanız \"Bu sürümü atla\" diyebilirsiniz; " +
                               "o sürüm için bir daha rahatsız edilmezsiniz. Elle kontrol " +
                               "ettiğinizde yine görünür.")
                }
            },

            new Adim
            {
                Bolum = "6 · AYARLAR VE BAKIM",
                Baslik = "Bir şey çalışmıyorsa",
                Ozet = "En sık karşılaşılan beş durum ve çözümleri.",
                Bloklar = new[]
                {
                    Blok.H("Kısayol tuşu hiçbir şey yapmıyor"),
                    Blok.P("Büyük ihtimalle başka bir program o tuşu kapmıştır ya da UIBUL'un ikinci " +
                           "bir kopyası açıktır. UIBUL tek örnek çalışır; ikinci kez açarsanız " +
                           "var olan pencere öne gelir. Alttaki konsolda kısayolların kayıt " +
                           "durumu yazar — oraya bakın."),

                    Blok.H("Yönetici olarak çalışan uygulamaları okuyamıyor"),
                    Blok.P("Windows'un güvenlik kuralıdır: düşük yetkili bir program yüksek yetkili " +
                           "olanı okuyamaz. UIBUL'a sağ tıklayıp \"Yönetici olarak çalıştır\" deyin."),

                    Blok.H("Element özellikleri eksik geliyor"),
                    Blok.P("Collection Profile'ı Full yapın ve F5 ile yenileyin. Bazı uygulamalar " +
                           "bilgiyi yalnız derin tarama sırasında verir."),

                    Blok.H("Chrome'da F11 tam ekran çalışmıyor"),
                    Blok.P("Beklenen davranış. F11'i UIBUL global olarak kaydettiği için tarayıcı " +
                           "tuşu hiç görmez. Ayar dosyasından AtlasKisayolu'nu \"Ctrl+F11\" yapın."),

                    Blok.H("Yakalama çok yavaş"),
                    Blok.P("Collection Profile'ı Standard veya Quick'e alın. Full profil " +
                           "her elementte 100'den fazla özelliği tek tek sorar."),

                    Blok.Ipucu("Her durumda ilk bakılacak yer alttaki konsoldur — program ne yaptığını " +
                               "ve nerede takıldığını oraya yazar.")
                }
            },

            new Adim
            {
                Bolum = "6 · AYARLAR VE BAKIM",
                Baslik = "Kısayol kartı — hepsi bir arada",
                Ozet = "Öğretici bitti. Bu sayfayı ekran görüntüsü alıp saklayabilirsiniz.",
                Bloklar = new[]
                {
                    Blok.T("F1", "İncelemeyi başlat (pencere gizlenir)"),
                    Blok.T("F2", "İncelemeyi durdur (pencere geri gelir)"),
                    Blok.T("F3", "İncelemeyi başlat (pencere görünür kalır)"),
                    Blok.T("F4", "Deklanşör — basılı tut, hedefe git, bırak"),
                    Blok.T("F5", "Seçili elementi yenile"),
                    Blok.T("F6", "TXT rapor → masaüstü + arşiv"),
                    Blok.T("F7", "Tam yakalama → masaüstü + arşiv"),
                    Blok.T("F8", "Tam yakalama → sadece arşiv"),
                    Blok.T("F9", "Bölge ekran görüntüsü"),
                    Blok.T("F10", "Son yakalama yolunu yapıştır"),
                    Blok.T("F11", "Kimlik şeritli kare (Atlas)"),
                    Blok.T("Ctrl+S", "Hızlı kaydet"),
                    Blok.T("Ctrl+C", "Element verisini kopyala"),
                    Blok.T("ESC", "Bölge seçimini iptal et"),
                    Blok.Ipucu("Bu öğreticiyi tekrar açmak için: üst şeritteki 🎓 ÖĞRETİCİ düğmesi " +
                               "ya da Help ▸ Öğretici. Daha ayrıntılı başvuru için 📖 KILAVUZ."),
                    Blok.Bilgi("İyi kullanımlar. Takıldığınız yerde alttaki konsola bakmayı unutmayın.")
                }
            }
        };
    }
}
