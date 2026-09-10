# Proje: UIBUL — Universal UI Element Inspector

> Windows uygulamalarının ve web sayfalarının arayüzünü **içeriden** okuyan
> inceleme aracı; aynı zamanda hızlı ekran görüntüsü ve arşivleme programı.
> Kullanıcı belgeleri: [`README.md`](README.md) · [`Docs/TANITIM.md`](Docs/TANITIM.md)

**Windows 10/11 · .NET 10 · WPF · Sürüm 3.1.0** (tek doğruluk kaynağı:
`UIElementInspector/UIElementInspector/UIElementInspector.csproj` `<Version>`)

---

## ClaudEmre

Bu proje 2026-09-10'da ClaudEmre sistemine kaydoldu.
Koordinatör oturumu `/claudemre-basla` ile açılır; gün sonunda
`/claudemre-bitir` çağrılır.

```
oturumlar/TESPIH.md      iş sırası — FAYDA ÷ EMEK ile dizili
oturumlar/CEPHANE.md     abonelik kademesi ve vites
oturumlar/<AD>.md        oturum şartnamesi (açılış prompt'u)
oturumlar/<AD>-ILERLEME.md   o oturumun ilerleme defteri
BEKLEYENLER.md           Emre'den beklenenler, kova etiketli
```

🔴 **Durum mesajla değil DOSYAYLA akar.** Her kalem bitince kendi ilerleme
dosyana yaz ve **pathspec'li** commit at:
```bash
git commit -F - -- oturumlar/<AD>-ILERLEME.md
```
Depo tek ağaçtır; `git add -A` **kullanılmaz** — başkasının satırını
commit'lersin.

Doktrin: `C:/Users/ana/Documents/Projects/ClaudEmre/` —
`CLAUDEMRE.md` · `YASALAR.md` · `KISALTMALAR.md` (⭐ `*mgy` `*yyy` `*iii` `*kii` `*ct`)

---

## Kod haritası

```
UIElementInspector/UIElementInspector/       ana uygulama (WPF)
  App.xaml.cs                                🔴 TEK ÖRNEK MUTEX'i burada
  MainWindow.xaml.cs                         ana pencere
  Core/Detectors/    5 algılama motoru — IElementDetector + UIAutomation ·
                     Win32 · MSHTML · WebView2 · Playwright
  Core/Models/       AppSettings · ElementInfo · CollectionProfile ·
                     ArchiveItem · InspectionSession
  Core/Utils/        ExportManager (JSON/TXT) · ArchiveManager ·
                     ScreenshotHelper · AtlasKare · AtlasDamgasi ·
                     SelectorGenerator · UpdateService · Logger · Bildirim
  Services/          HotkeyService (global kısayol) · MouseHookService
  Windows/           öğretici · kılavuz · ayarlar · bölge seçici · güncelleme
Installer/UibulSetup/    kurulum sihirbazı (WinForms)
tools/build-release.ps1  tek komutla setup üretme
Docs/                    kullanıcıya giden belgeler
```

Ölçüldü (2026-09-10): **87 `.cs` · 31.281 satır · 39 commit** (ilk: 2025-11-23).

---

## 🔴 Bu kod tabanının BİLİNEN TUZAKLARI — ölçüldü, uydurulmadı

### ① Tek örnek mutex'i — ikinci örnek KENDİNİ KAPATIR
`App.OnStartup` `Global\UIElementInspector_TekOrnek` mutex'ini alır; alamayan
örnek `Shutdown()` çağırır ve **hiçbir çıktı üretmeden ölür**. Sebebi
yazılı: global kısayolu (`RegisterHotKey`) yalnız ilk örnek alabilir.
⇒ Yeni bir çalıştırma kipi (komut satırı dâhil) eklerken bu kapı **ilk
düşünülecek şeydir**, sonuncu değil.

### ② `OutputType=WinExe` — konsol YOK
`csproj` `WinExe` üretir. `Console.WriteLine` bir yere gitmez; standart
çıkışa yazmak için `AttachConsole(ATTACH_PARENT_PROCESS)` gerekir.

### ③ Global kısayol bedeli ORTAMA yazılır
Global tuş, UIBUL açıkken **başka programlardan çalınır**. 2026-09-07'de
12 tuş 4'e indirildi; işlevler silinmedi, düğmelere taşındı.
Ders: `ClaudEmre/yasalar/gelen/2026-09-07-ana-global-kaynagi-rezerve-etmenin-bedelini-ortam-oder.md`

### ④ Varsayılan değiştirmek, diskteki ayarı değiştirmez
Atlas tuşu kodda `F11 → F4` oldu ama kullanıcının ayar dosyasında `F11`
duruyordu; `AppSettings.Load` bir **göç adımı** taşıyor. Bir varsayılanı
değiştiren, saklanmış nüshaları da göçürür — ve sınavı **iki yönlüdür**
(değeri OLAN ve OLMAYAN kullanıcı).
Ders: `ClaudEmre/yasalar/gelen/2026-09-07-ana-yeni-varsayilan-diskteki-eski-degeri-degistirmez.md`

---

## Derleme

```powershell
dotnet build UIElementInspector\UIElementInspector\UIElementInspector.csproj
dotnet run   --project UIElementInspector\UIElementInspector\UIElementInspector.csproj
powershell -ExecutionPolicy Bypass -File tools\build-release.ps1   # setup .exe
```
⚠️ `bin/`, `obj/`, `publish/`, `Logs/` **gitignore'da** — derleme çıktısı
commit'lenmez.
