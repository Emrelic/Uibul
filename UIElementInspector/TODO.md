# UIElementInspector - TODO List

## ✅ Tamamlanan Özellikler

### Çekirdek Özellikler
- [x] WPF masaüstü uygulaması oluşturuldu
- [x] Modüler detector mimarisi (IElementDetector interface)
- [x] Kapsamlı ElementInfo modeli (100+ özellik)
- [x] Multiple collection profiles (Quick, Standard, Full, Custom)

### Algılama Teknolojileri
- [x] **UI Automation** - Windows masaüstü uygulamaları için
- [x] **WebView2/CDP** - Modern Chromium tarayıcılar için (Chrome, Edge)
- [x] **MSHTML/IHTMLDocument** - Internet Explorer ve eski tarayıcılar için
- [ ] **Playwright** - Kod hazır ama paket yüklenemedi (ağ problemi)

### UI Özellikleri
- [x] 3 panelli layout (TreeView, Properties, Console)
- [x] Mouse hover ile element algılama
- [x] Click mode ile element yakalama
- [x] Tüm window elementlerini toplama
- [x] Raw ve kategorize property görünümleri
- [x] Element arama fonksiyonu
- [x] Keyboard shortcuts (F1, ESC, F5, Ctrl+S)

### Export Özellikleri
- [x] CSV export
- [x] TXT export (inspect.exe formatı)
- [x] JSON export
- [x] XML export
- [x] HTML export (interaktif tablo)

### Screenshot Özellikleri
- [x] Element screenshot
- [x] Bölge screenshot
- [x] Tam ekran screenshot
- [x] Highlight özelliği

## 🚧 Eksik Kalan Özellikler

### Yüksek Öncelik
1. **Region Selector (Sürükle-Bırak)**
   - Overlay window oluşturulması gerekiyor
   - Mouse ile dikdörtgen çizme
   - Seçilen bölgedeki tüm elementleri toplama
   - Dosya: `Services/RegionSelectorService.cs` oluşturulmalı

2. **XPath ve CSS Selector Generator**
   - Daha akıllı selector üretimi
   - Multiple selector stratejileri
   - Selector doğrulama
   - Dosya: `Core/Utils/SelectorGenerator.cs` oluşturulmalı

### Orta Öncelik
3. **Playwright Integration**
   - Paket yükleme sorunu çözülmeli
   - PlaywrightDetector.cs implementasyonu tamamlanmalı
   - Cross-browser desteği eklenecek

4. **LegacyIAccessiblePattern Desteği**
   - UIAutomationDetector.cs içinde yorum satırında
   - Eski uygulamalar için önemli
   - System.Windows.Automation.LegacyIAccessiblePattern referansı eklenmeli

5. **MSHTML Checkbox/Radio Desteği**
   - MSHTMLDetector.cs line 318-330 arası yorum satırında
   - Dynamic property access sorunu çözülmeli

### Düşük Öncelik
6. **Import/Export Session**
   - Session kaydetme ve yükleme
   - JSON formatında session dosyaları

7. **Settings Penceresi**
   - Collection profile ayarları
   - Hotkey özelleştirme
   - Export varsayılan ayarları

8. **Element Tree Building**
   - WebView2Detector.GetElementTree() implementasyonu
   - MSHTMLDetector.GetElementTree() implementasyonu

9. **Color Picker Tool**
   - Element rengini alma
   - RGB/HEX değerleri

10. **Performance Optimizasyonu**
    - Large element collection için pagination
    - Memory usage optimizasyonu
    - Async operation iyileştirmeleri

## 📝 Notlar

### Bilinen Sorunlar
- Playwright paketi yüklenemiyor (network timeout)
- WebView2 initialization async olduğu için ilk başta null olabilir
- MSHTML dynamic property access hataları var

### Geliştirme Ortamı
- .NET 8.0
- WPF
- Windows 10/11
- Visual Studio 2022 veya VS Code önerilir

### Test Edilmesi Gerekenler
- Windows 7/8 uyumluluğu
- Yüksek DPI ekran desteği
- Multi-monitor desteği
- UAC yükseltilmiş uygulamalar

## 🔧 Nasıl Devam Edilir

1. Bu TODO listesindeki özellikleri sırayla implement edin
2. Her özellik için ayrı branch oluşturun
3. Test yazın (birim testler eksik)
4. Documentation güncelleyin

## 📦 Eksik NuGet Paketleri
```xml
<!-- Playwright için (ağ sorunu çözülünce) -->
<PackageReference Include="Microsoft.Playwright" Version="1.40.0" />
```

## 📚 Yararlı Kaynaklar
- [UI Automation Documentation](https://docs.microsoft.com/en-us/dotnet/framework/ui-automation/)
- [WebView2 Documentation](https://docs.microsoft.com/en-us/microsoft-edge/webview2/)
- [MSHTML Reference](https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752041(v=vs.85))

---
Son Güncelleme: 2024-11-23
Proje Durumu: %95 Tamamlandı - Ana özellikler çalışıyor, opsiyonel özellikler eksik