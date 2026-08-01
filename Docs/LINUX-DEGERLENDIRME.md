# Linux sürümü — fizibilite değerlendirmesi

**Kısa cevap:** Evet, üretilebilir — ama bu bir *taşıma* değil, **kısmi yeniden yazım**dır.
Kodun yaklaşık **%35'i olduğu gibi taşınır**, geri kalanı yeniden yazılmalıdır. Ve
sonuçta ortaya çıkan program, Windows sürümüyle **aynı yetenekte olmaz** — sebebi
benim tercihim değil, Linux masaüstünün kendi yapısıdır.

Aşağıda ölçtüğüm sayılar ve dürüst bir yol haritası var.

---

## 1. Mevcut kodun Windows'a bağlılığı (ölçüldü)

| Bağımlılık | Kaç dosyada | Linux'ta durumu |
|---|---:|---|
| `System.Windows` (WPF) | 81 | ❌ Yok — Avalonia'ya yeniden yazım |
| `System.Windows.Forms` | 55 | ❌ Yok — Avalonia'ya yeniden yazım |
| `System.Windows.Automation` (UIA) | 32 | ❌ Yok — AT-SPI2 ile yeniden yazım |
| `System.Drawing` | 27 | ⚠️ Kısmen — SkiaSharp/ImageSharp'a geçmeli |
| P/Invoke (`user32`, `gdi32`, …) | **74 çağrı / 9 dosya** | ❌ Tamamı yeniden yazım |
| `mshtml` (Internet Explorer) | 1 | ❌ Linux'ta IE yok — **tamamen düşer** |
| `Microsoft.Web.WebView2` | 1 | ⚠️ CDP'ye çevrilmeli (protokol aynı, taşınabilir) |
| `Microsoft.Playwright` | 1 | ✅ Zaten çapraz platform |

Toplam kod: ~23.000 satır.

---

## 2. Ne taşınır, ne yeniden yazılır?

### ✅ Olduğu gibi taşınır (~%35)

Bu dosyalar saf .NET; hiçbir değişiklik gerektirmez:

- `Core/Models/` — `ElementInfo`, `InspectionSession`, `CollectionProfile`, `AppSettings`
- `Core/Utils/ExportManager.cs` — CSV / JSON / XML / HTML / TXT üretimi
- `Core/Utils/SelectorGenerator.cs` — XPath ve CSS seçici üretimi (saf mantık)
- `Core/Utils/Logger.cs`, `ArchiveManager.cs` — yol ayıraçları düzeltilerek
- `Core/Utils/UpdateService.cs` — `HttpClient` tabanlı, zaten çapraz platform
- `Windows/TutorialContent.cs` — öğretici metinleri (veri, kod değil)

### 🔁 Yeniden yazılır (~%65)

| Katman | Windows'ta | Linux'ta karşılığı | Zorluk |
|---|---|---|---|
| Arayüz | WPF + WinForms | **Avalonia UI** (XAML, en yakın akraba) | Orta — mekanik ama uzun |
| Element okuma | UI Automation | **AT-SPI2** (D-Bus üzerinden) | Yüksek |
| Pencere sorgulama | Win32 API | X11: `XQueryTree` / Wayland: yok | Yüksek |
| Ekran görüntüsü | GDI+ `BitBlt` | X11: `XGetImage` / Wayland: portal | **Çok yüksek (Wayland)** |
| Global kısayol | `RegisterHotKey` | X11: `XGrabKey` / Wayland: portal | **Çok yüksek (Wayland)** |
| Tarayıcı içi okuma | WebView2 | **CDP doğrudan** (`--remote-debugging-port`) | Düşük — protokol aynı |
| Pano (resim + yol) | Win32 clipboard | X11 selection / `wl-clipboard` | Orta |
| Eski IE motoru | MSHTML | — | Düşer |

---

## 3. Asıl mesele: X11 mi, Wayland mı?

Bu, projenin kaderini belirleyen tek soru. Modern Ubuntu, Fedora ve diğerleri artık
**varsayılan olarak Wayland** kullanıyor.

### X11'de (eski ama hâlâ yaygın)

Windows sürümünün neredeyse tamamı karşılanabilir:
- ✅ Global kısayollar — `XGrabKey` ile çalışır
- ✅ Ekran görüntüsü — `XGetImage` ile, izin sormadan
- ✅ Başka pencerelerin ağacını okuma — mümkün
- ✅ Fare konumundaki elementi bulma — mümkün

### Wayland'da (yeni, varsayılan)

Wayland **tasarım gereği** bir uygulamanın başka uygulamaları gözetlemesini engeller.
Bu bir eksiklik değil, bilinçli bir güvenlik kararıdır — ve bizim aracımızın yaptığı
şey tam olarak "başka uygulamaları gözetlemek"tir.

- ⚠️ **Global kısayol:** Standart yok. `xdg-desktop-portal` içindeki `GlobalShortcuts`
  arayüzü çok yeni ve masaüstü ortamlarına göre desteği değişken. F1–F11'in her
  uygulamada çalışması **garanti edilemez**.
- ⚠️ **Ekran görüntüsü:** Yalnızca portal üzerinden ve **kullanıcı her oturumda izin
  penceresi onaylayarak**. "F9'a bas, anında kaydet" akışı Wayland'da bozulur.
- ✅ **Element okuma:** AT-SPI2 çalışır (D-Bus üzerinden, ekran sunucusundan bağımsız).
  Bu iyi haber — aracın *asıl* işi Wayland'da da yapılabilir.

**Sonuç:** Linux sürümünün *element inceleme* tarafı iyi çalışır; *ekran görüntüsü ve
global kısayol* tarafı Wayland'da belirgin biçimde zayıflar.

---

## 4. AT-SPI2 ne kadar iyi?

AT-SPI2, Linux'un erişilebilirlik katmanıdır — Windows'taki UI Automation'ın karşılığı.

| Uygulama türü | AT-SPI desteği |
|---|---|
| GTK (GNOME, Nautilus, GIMP…) | ✅ Çok iyi |
| Qt / KDE | ✅ İyi (`QT_ACCESSIBILITY=1` gerekebilir) |
| Firefox | ✅ İyi |
| Chrome / Chromium / Electron | ⚠️ `--force-renderer-accessibility` bayrağı gerekir |
| Java (Swing) | ⚠️ Ek köprü gerekir |
| Flutter, oyunlar, canvas editörler | ❌ Hiç bildirmez |

Yani Windows'takine benzer bir "her şeyi okur" iddiası Linux'ta kurulamaz.
Chromium tabanlı uygulamalar için CDP yolu (tarayıcıyı hata ayıklama portuyla
başlatmak) daha güvenilir olur.

---

## 5. Öneri: üç kademeli yol

### Kademe A — "Günlük araç" (küçük, hızlı, kesin çalışır)

Yalnızca ekran görüntüsü tarafı: bölge seçimi, otomatik kayıt, panoya kopyalama,
kimlik şeritli kare (F11 mantığı), arşiv, dışa aktarma.

- Avalonia arayüz + X11/portal ekran yakalama
- Element okuma **yok**
- Tahmini: ~2.000 satır, orta büyüklükte bir iş
- Wayland'da bile kullanılabilir (portal izniyle)

### Kademe B — "İnceleme aracı" (asıl değer)

Kademe A + AT-SPI2 ile element okuma + CDP ile tarayıcı okuma + dışa aktarma.

- Windows sürümünün *yazılımcıya faydası* olan kısmının çoğu gelir
- Tahmini: ~6.000–8.000 satır
- GTK/Qt/Firefox uygulamalarında iyi, Electron'da bayrakla, bazı uygulamalarda hiç

### Kademe C — "Tam eşitlik"

Gerçekçi değil. Wayland'ın izin modeli ve AT-SPI'nin kapsama boşlukları yüzünden
Windows sürümüyle bire bir aynı davranış **üretilemez**. Bunu vaat etmem doğru olmaz.

---

## 6. Benim yapamayacağım şey: doğrulama

Açıkça söylemem gereken bir sınır var: **bu makine Windows.** Linux kodunu yazabilirim,
ama burada **derleyemem ve çalıştıramam**. X11/Wayland/AT-SPI davranışı ancak gerçek bir
Linux masaüstünde ölçülebilir.

Yani Linux sürümü için üç seçenekten biri gerekir:

1. Sizde bir Linux makinesi/sanal makine olması (kodu yazarım, siz çalıştırıp
   çıktıyı bana verirsiniz — böyle çalışabiliriz)
2. WSL2 + WSLg (kısmen işe yarar; ama WSLg Wayland tabanlıdır ve gerçek masaüstü
   davranışını tam yansıtmaz)
3. Kodu "en iyi tahminle" yazmam ve doğrulamayı sizin yapmanız — hata payı yüksek olur

---

## 7. Karar için sorular

Devam etmek isterseniz şunları bilmem gerekiyor:

1. **Hangi kademe?** A (günlük araç), B (inceleme aracı) yoksa önce A sonra B mi?
2. **Hangi dağıtım ve oturum?** Ubuntu 24.04 + Wayland mı, X11 mi? Bu, mimariyi
   doğrudan belirliyor.
3. **Nasıl dağıtılacak?** `.deb` paketi mi, AppImage mı, Flatpak mı?
   (AppImage en kolayı: tek dosya, kurulum gerektirmez — Windows'taki setup.exe'nin karşılığı)
4. **Test imkânı var mı?** Yukarıdaki 3 seçenekten hangisi?

---

## 8. Bu arada: Windows sürümü hazır

Linux kararı beklerken Windows tarafı tamamlanmış durumda — setup, otomatik
güncelleme, öğretici ve belgeler çalışıyor. Linux sürümü ayrı bir dal
(`linux` branch) olarak geliştirilebilir; ortak kod (`Core/Models`, `ExportManager`,
`SelectorGenerator`, `UpdateService`) paylaşılan bir `UIBUL.Core` projesine
çıkarılırsa iki sürüm birlikte yürüyebilir.
