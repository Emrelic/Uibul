# UIBUL — Universal UI Element Inspector

> Ekranda gördüğünüz her şeyi okuyan, yakalayan ve belgeleyen araç.
> Windows 10/11 · .NET 10 · WPF

Bu belgenin biçimli ve yazdırılabilir hâli: [`TANITIM.html`](TANITIM.html)
(kurulumla birlikte gelir, program içinden **Help ▸ Documentation** ile açılır).

---

## Bir cümlede

UIBUL, Windows'ta çalışan **her uygulamanın ve her web sayfasının arayüzünü içeriden
okuyabilen** bir inceleme aracıdır; aynı zamanda gündelik işler için hızlı bir
**ekran görüntüsü ve arşivleme** programıdır.

İkisi birden, çünkü aynı altyapıyı kullanıyorlar: bir düğmenin kimliğini okumak için
ekranı ve pencere ağacını çözmek gerekir, ekranın bir bölgesini kesip kaydetmek de
aynı yeteneğin daha basit hâlidir.

---

## Ne yapar?

| | |
|---|---|
| **Element okur** | Ad, tip, kimlik, konum, HTML, erişilebilirlik — 100+ özellik, 10 kategoride |
| **Beş motorla** | UI Automation · WebView2/CDP · MSHTML · Win32 · Playwright (deneysel) |
| **Yakalar ve paketler** | Tek tuşla element analizi + DOM + kaynak kod + 3 ekran görüntüsü |
| **Ekran görüntüsü alır** | Bölge seçimi, otomatik kayıt, panoya hem resim hem yol |
| **Dışa aktarır** | CSV · JSON · XML · HTML · TXT |
| **Arşivler** | Tarihli, okunabilir adlarla; program içinden yeniden bulunur |

---

## Bir yazılımcının ne işine yarar?

### 1. Test otomasyonunda selector bulma — en somut fayda

Selenium, Playwright, Appium veya WinAppDriver ile test yazarken zamanın büyük kısmı
koda değil, **elementleri güvenilir biçimde bulmaya** gider. Web tarafında DevTools iş
görür; ama masaüstü uygulamalarında elinizde genellikle hiçbir şey yoktur —
Microsoft'un `inspect.exe`'si Windows SDK ile gelir, arayüzü ilkeldir, çıktı alma
imkânı yok denecek kadar azdır.

UIBUL'un kazandırdığı net: **web ve masaüstü için tek araç**, **hazır XPath/CSS
seçicileri**, **AutomationId + pattern listesi**, ve en önemlisi **toplu dışa aktarma**.
Bir ekranın tüm elementlerini gezip JSON'a döküp Page Object sınıfına dönüştürebilirsiniz.
`inspect.exe` ile aynı iş elle kopyalayarak saatler alır.

### 2. Hata raporunu tartışmasız hâle getirmek

"Bende tekrar etmiyor", yazılım ekiplerinde en çok zaman yakan cümledir. `F7` tek tuşta
ekran görüntüsü + element özellikleri + DOM + pencere bilgisini aynı klasöre koyar.
Klasörü issue'ya eklemek, üç ekran görüntüsü ve bir paragraf açıklamadan fazla bilgi taşır.

### 3. Erişilebilirlik denetimi

Ekran okuyucuların gördüğü katman ile UIBUL'un okuduğu katman **aynıdır** (UI Automation /
ARIA). Adı boş, rolü yanlış veya etiketsiz elementleri gösterir. Full profil + HTML export,
ücretli a11y araçlarının temel taramasına yakın sonuç verir.

### 4. Belgesiz ve eski (legacy) sistemleri çözmek

Kaynak kodu kaybolmuş bir iç uygulamayı otomatikleştirmeniz gerekiyorsa UIBUL iç yapıyı
dışarıdan haritalandırır: pencere sınıfları, kontrol hiyerarşisi, handle'lar. Bu harita
olmadan böyle bir işe başlamak körlemedir.

### 5. Yapay zekâya ekran bağlamı vermek

Bir dil modelinden "şu ekrandaki düğmeye tıklayan kodu yaz" derken en büyük eksik, modelin
ekranı görmemesidir. `F7` çıktısı bu boşluğu doldurur: görüntü + element özellikleri +
seçiciler birlikte gider.

### Dürüst sınırlar

- UIBUL bir **inceleme** aracıdır, test koşucusu değil. Selector verir; testi siz yazarsınız.
- Yönetici olarak çalışan uygulamaları okumak için kendisinin de yönetici olması gerekir
  (Windows'un güvenlik kuralı, aşılamaz).
- Özel çizimli arayüzler (oyunlar, Skia/Flutter ile çizilmiş bazı uygulamalar, canvas
  tabanlı editörler) işletim sistemine element bildirmez; orada yalnızca ekran görüntüsü
  aracı olarak kalır.
- Playwright motoru deneyseldir ve ayrıca tarayıcı indirmek ister.

---

## Yazılımcı olmayanlar için

En çok kullanılan tuş, yazılımla hiç ilgisi olmayan **`F9`**'dur. Windows'un Ekran
Alıntısı aracına benzer, üç farkla: dosya **otomatik kaydedilir**, panoya **hem resim
hem dosya yolu** konur, ve **her uygulamada** çalışır.

| Senaryo | Nasıl |
|---|---|
| Destek talebi | Hata mesajını `F9` ile seçin, sohbete `Ctrl+V` |
| Fatura / dekont / bilet arşivi | Onay ekranını `F9`; dosya adı tarihli olur |
| Alışverişte fiyat kanıtı | İndirimli fiyatı `F9` ile kaydedin |
| Ders / toplantı notu | Grafiği veya slaytı `F9` ile alıp not defterine yapıştırın |
| Uzun form güvenlik ağı | Doldurduğunuz bölümleri `F9` ile kaydedin |
| Dosyayı birine göndermek | `F9` → sohbete geç → `F10` (yol yapıştırılır) |
| Tarihli, kimlikli kanıt | `F11` — tarih ve konum görüntünün **içine** basılır |

---

## Kısayol tuşları

| Tuş | Ne yapar | Ne zaman |
|---|---|---|
| `F1` | İncelemeyi başlat — pencere gizlenir | Tek monitörde |
| `F2` | İncelemeyi durdur | Her zaman; global |
| `F3` | İncelemeyi başlat — pencere görünür kalır | Çift monitörde |
| `F4` | Deklanşör: basılı tut → hedefe git → bırak | Açılır menü, dropdown, tooltip |
| `F5` | Seçili elementi yeniden analiz et | Dinamik içerikte |
| `F6` | TXT rapor → masaüstü + arşiv | Sadece bilgi lazımsa |
| `F7` | Tam yakalama → masaüstü + arşiv | Hata raporu (en kapsamlı) |
| `F8` | Tam yakalama → sadece arşiv | Uzun seanslarda |
| `F9` | Bölge ekran görüntüsü | Gündelik hayatta en çok |
| `F10` | Son yakalama yolunu yapıştır | Dosya gönderirken |
| `F11` | Kimlik şeritli kare | Tarihli kanıt gerektiğinde |
| `Ctrl+S` | Hızlı kaydet | Acele varken |
| `Ctrl+C` | Element verisini kopyala | — |
| `ESC` | Bölge seçimini iptal et | — |

> ⚠️ Bu tuşlar **global** kaydedilir: UIBUL açıkken her uygulamada çalışırlar. Bedeli,
> aynı tuşu kullanan başka programların o tuşu görmemesidir — örneğin **Chrome'un `F11`
> tam ekranı UIBUL açıkken çalışmaz**. Rahatsız ederse `settings.json` içinde
> `AtlasKisayolu` → `"Ctrl+F11"` yapın.

---

## Kurulum

1. `UIBUL_Setup.exe` dosyasına çift tıklayın
2. Kurulum bilgisayarı denetler (Windows sürümü, disk, WebView2); eksik varsa kurmayı teklif eder
3. İleri → İleri → Kur — **yönetici hakkı sormaz**
4. Program ilk açılışta adım adım öğreticiyi gösterir

**Ön koşul yok.** Kurulum dosyası .NET çalışma zamanını da içerir; hedef bilgisayarda
hiçbir şeyin kurulu olması gerekmez. Tek isteğe bağlı bileşen WebView2 Runtime'dır
(tarayıcı içi element okuma için); yoksa program yine çalışır, o motor devre dışı kalır.

**Kaldırma:** Ayarlar ▸ Uygulamalar ▸ UIBUL. Ayarlarınız ve arşiviniz
`%AppData%\UIElementInspector` altında **silinmeden kalır**.

---

## Güncellemeler

UIBUL yeni sürümleri kendisi arar (açılışta, günde en fazla bir kez). Yeni sürüm varsa
pencere açılır; yoksa hiçbir şey olmaz. "Şimdi güncelle" → indirilir, program kapanır,
güncellenir, yeniden açılır. Ayarlar ve arşiv korunur.

- Elle kontrol: **Help ▸ Güncellemeleri kontrol et**
- Bir sürümü istemiyorsanız: **Bu sürümü atla**
- Kapatmak için: `settings.json` → `OtomatikGuncellemeKontrolu: false`

İnternet yoksa açılışta **hata penceresi çıkmaz**; durum yalnızca konsola yazılır.

---

## Sorun giderme

| Durum | Sebep / çözüm |
|---|---|
| Kısayol çalışmıyor | Başka program tuşu kapmış olabilir; alttaki konsolda kayıt durumu yazar |
| Yönetici uygulamaları okunmuyor | UIBUL'u "Yönetici olarak çalıştır" ile açın |
| Özellikler eksik | Collection Profile → **Full**, sonra `F5` |
| Chrome'da F11 çalışmıyor | Beklenen; yukarıdaki kısayol uyarısına bakın |
| Yakalama yavaş | Collection Profile → **Standard** veya **Quick** |
| Masaüstü doldu | Tools ▸ Settings ▸ çıktı klasörü; ya da `F8` kullanın |

---

## Teknik künye

| | |
|---|---|
| Platform | Windows 10 / 11, 64-bit |
| Teknoloji | .NET 10 · WPF · C# |
| Motorlar | UI Automation, WebView2/CDP, MSHTML, Win32, Playwright (deneysel) |
| Dışa aktarma | CSV, JSON, XML, HTML, TXT |
| Ayarlar | `%AppData%\UIElementInspector\settings.json` |
| Günlükler | `%AppData%\UIElementInspector\Logs` |
| Ön koşul | Yok (self-contained) |
| Depo | https://github.com/Emrelic/Uibul |

© 2026 Emrelic
