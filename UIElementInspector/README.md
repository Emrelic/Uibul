# Universal UI Element Inspector

Windows uygulamaları ve web tarayıcıları için kapsamlı bir UI element inceleme aracı.

## 🎯 Özellikler

- **Çoklu Algılama Teknolojileri**: UI Automation, WebView2/CDP, MSHTML
- **Evrensel Destek**: Masaüstü uygulamaları, modern tarayıcılar (Chrome, Edge), Internet Explorer
- **100+ Element Özelliği**: Tüm UI element özelliklerini toplar
- **Screenshot Desteği**: Element, bölge veya tam ekran görüntüleri
- **Çoklu Export Formatları**: CSV, TXT, JSON, XML, HTML
- **Gerçek Zamanlı İnceleme**: Mouse hover ile anında element tespiti

## 🚀 Hızlı Başlangıç

### Gereksinimler
- Windows 10/11
- .NET 10.0 Runtime (kurulum paketinde gömülü gelir)
- Visual Studio 2022 (geliştirme için)

### Kurulum ve Çalıştırma

```bash
# Projeyi derle
cd UIElementInspector/UIElementInspector
dotnet build

# Uygulamayı çalıştır
dotnet run
```

## 📖 Kullanım Kılavuzu

### Temel Kullanım

1. **İncelemeyi Başlat**: F1 tuşuna basın veya "Start Inspection" butonuna tıklayın
2. **Mod Seçin**: Hover, Click, Region veya Full Window modlarından birini seçin
3. **Element Algıla**: Mouse'u herhangi bir UI elementi üzerine getirin
4. **Özellikleri İncele**: Sağ panelde detaylı özellikleri görün
5. **Veri Export**: File > Export menüsünden istediğiniz formatta kaydedin

### Klavye Kısayolları

- **F1** - İncelemeyi başlat
- **ESC** - İncelemeyi durdur
- **F5** - Seçili elementi yenile
- **F9** - Ekran bölgesi görüntüsü
- **F11** - Tarih Atlası karesi (aşağıya bakın)
- **Ctrl+S** - Hızlı export
- **Ctrl+C** - Element verilerini kopyala

### F11 — Tarih Atlası karesi

Atlas sitesindeki bir kusuru bildirmek için kullanılır. Bölge seçilir; kırpılan
alanın etrafına 3 px kırmızı çerçeve, altına da ince bir kimlik şeridi çizilir:

```
1361-02-01 · 41.35N 26.50E · z6 · Osmanlı Tarih Atlası
Madde: Edirne'nin fethi
```

Bu satırlar görüntünün **içine** basılır; dosya adı kaybolsa bile bilgi kareyle
birlikte gider. Aynı metin panoya da konur. Kare hem panoya kopyalanır hem
`%LOCALAPPDATA%\TarihAtlasiKare` altına PNG olarak kaydedilir (OneDrive dışında;
son 50 kare tutulur, eskiler silinir).

**Bilgi nereden okunur:** tarayıcı penceresinin başlığından. Atlas sayfası
`document.title`'ı şu hâle getirmelidir:

```
Osmanlı Tarih Atlası · 1361-02-01 · 41.35N 26.50E · z6 · <açık kronoloji maddesi>
```

**Damga yoksa ne olur:** kare yine alınır, ama şerit koyu kırmızı zeminle
`TARİH/KOORDİNAT OKUNAMADI` der ve karenin alındığı saati yazar — tarih
**uydurulmaz**. Dosya adı `damgasiz_YYYY-AA-GG_SSDDss.png` olur. Ayrıca her
kareden sonra ekranın sağ altında kısa bir bildirim çıkar; ana pencere gizliyken
de ne olduğu görünür.

Uygulama **tek örnek** çalışır: ikinci kez açılırsa var olan pencere öne gelir.
Sebebi kısayol: global kısayolu yalnız ilk örnek alabilir, ikincisi tuşu yutmayan
klavye kancasına düşer ve aynı tuşa iki örnek birden tepki verirdi.

Ayarlar (`%AppData%\UIElementInspector\settings.json`):
`AtlasKisayolu` · `AtlasKlasoru` · `AtlasEnUzunKenar` (varsayılan 1200) ·
`AtlasSonKareSayisi` (varsayılan 50) · `AtlasDamgasizIzin` (varsayılan true).

> ⚠️ **Ölçülmüş uyarı:** F11 global kısayol olarak kaydedildiği için, bu araç
> açıkken Chrome'un F11 tam ekran kısayolu **çalışmaz** (tuşu Windows bize
> yönlendiriyor, tarayıcı hiç görmüyor). İstemezseniz `AtlasKisayolu` değerini
> `"Ctrl+F11"` yapın — o kombinasyon boş ve çakışmasız.
>
> ⚠️ Görüntü maliyeti yalnız **piksel sayısına** bağlıdır; PNG/JPEG seçimi ve
> sıkıştırma kalitesi hiçbir şey değiştirmez. Bu yüzden kalite ayarı yoktur,
> tek gerçek tasarruf küçültmedir (en uzun kenar 1200 px'e indirilir; küçük
> kareler asla büyütülmez).

### İnceleme Modları

1. **Hover Mode**: Mouse üzerinde olduğu elementi gerçek zamanlı algılar
2. **Click Mode**: Tıklanan elementi yakalar ve incelemeyi durdurur
3. **Region Mode**: Dikdörtgen bölge seçerek elementleri toplar (geliştirme aşamasında)
4. **Full Window Mode**: Aktif penceredeki tüm elementleri toplar

### Collection Profilleri

- **Quick**: Temel özellikler, hızlı toplama (< 1 saniye)
- **Standard**: Standart özellikler, orta hız (1-3 saniye)
- **Full**: Tüm özellikler, detaylı toplama (3-10 saniye)
- **Custom**: Özelleştirilebilir profil

## 🏗️ Proje Yapısı

```
UIElementInspector/
├── Core/
│   ├── Detectors/          # Algılama teknolojileri
│   │   ├── UIAutomationDetector.cs
│   │   ├── WebView2Detector.cs
│   │   └── MSHTMLDetector.cs
│   ├── Models/
│   │   └── ElementInfo.cs  # Element veri modeli
│   └── Utils/
│       ├── ExportManager.cs    # Export işlemleri
│       └── ScreenshotHelper.cs # Screenshot işlemleri
├── Services/
│   ├── MouseHookService.cs     # Global mouse hook
│   └── HotkeyService.cs       # Klavye kısayolları
├── MainWindow.xaml             # Ana pencere UI
└── MainWindow.xaml.cs         # Ana pencere logic
```

## 🔧 Teknik Detaylar

### Desteklenen Teknolojiler

| Teknoloji | Kullanım Alanı | Durum |
|-----------|---------------|--------|
| UI Automation | Windows masaüstü uygulamaları | ✅ Aktif |
| WebView2/CDP | Chrome, Edge tarayıcıları | ✅ Aktif |
| MSHTML | Internet Explorer, eski tarayıcılar | ✅ Aktif |
| Playwright | Cross-browser test otomasyonu | ⏳ Beklemede |

### Element Özellikleri

- **Temel**: Name, Type, Value, Class, ID
- **UI Automation**: AutomationId, ControlType, Patterns, States
- **Web**: TagName, innerHTML, href, XPath, CSS Selector
- **Pozisyon**: X, Y, Width, Height, BoundingRectangle
- **Hiyerarşi**: Parent, Children, TreeLevel
- **Erişilebilirlik**: ARIA attributes, Role, Label

## 📊 Export Formatları

| Format | Açıklama | Kullanım |
|--------|----------|----------|
| CSV | Virgülle ayrılmış değerler | Excel, veri analizi |
| TXT | inspect.exe formatı | Metin editörler, loglama |
| JSON | Yapılandırılmış veri | API entegrasyonu, programatik işleme |
| XML | Hiyerarşik veri | Kurumsal sistemler |
| HTML | İnteraktif tablo | Web görüntüleme, filtreleme |

## 🐛 Bilinen Sorunlar ve Çözümler

1. **Playwright paketi yüklenemiyor**
   - Sebep: Network timeout
   - Çözüm: VPN kapatın veya farklı network deneyin

2. **WebView2 başlangıçta null**
   - Sebep: Async initialization
   - Çözüm: Birkaç saniye bekleyin

3. **Yükseltilmiş uygulamalar algılanmıyor**
   - Sebep: UAC kısıtlamaları
   - Çözüm: Uygulamayı yönetici olarak çalıştırın

## 📝 Eksik Özellikler

Detaylı liste için [TODO.md](TODO.md) dosyasına bakın.

## 🤝 Katkıda Bulunma

1. Projeyi fork edin
2. Feature branch oluşturun (`git checkout -b feature/AmazingFeature`)
3. Değişikliklerinizi commit edin (`git commit -m 'Add some AmazingFeature'`)
4. Branch'e push edin (`git push origin feature/AmazingFeature`)
5. Pull Request açın

## 📄 Lisans

Bu proje eğitim ve geliştirme amaçlı oluşturulmuştur.

## 📞 İletişim

Sorularınız için GitHub Issues kullanabilirsiniz.

---
**Versiyon**: 3.1.0
**Durum**: Aktif Geliştirme (%95 Tamamlandı)
**Son Güncelleme**: 2026-08-01