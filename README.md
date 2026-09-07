# UIBUL — Universal UI Element Inspector

Windows uygulamalarının ve web sayfalarının arayüzünü içeriden okuyan inceleme aracı;
aynı zamanda gündelik işler için hızlı bir ekran görüntüsü ve arşivleme programı.

**Windows 10/11 · .NET 10 · WPF · Sürüm 3.1.0**

---

## Kurulum

[Releases](https://github.com/Emrelic/Uibul/releases/latest) sayfasından
`UIBUL_Setup.exe` indirin ve çalıştırın.

Ön koşul yoktur — .NET dâhil her şey kurulum dosyasının içindedir. Yönetici hakkı
sormaz. Program yeni sürümleri kendisi kontrol eder ve tek tıkla günceller.

> Windows SmartScreen uyarısı çıkarsa: **Ek bilgi ▸ Yine de çalıştır**.
> Kurulum dosyası dijital imzalı değildir (kod imzalama sertifikası ücretlidir).

---

## Belgeler

| Belge | Kimin için |
|---|---|
| [Tanıtım ve kullanım](Docs/TANITIM.md) | Kullanıcılar — ne yapar, nasıl kullanılır, yazılımcıya faydası |
| [Sürüm çıkarma](Docs/SURUM-CIKARMA.md) | Geliştirici — yeni sürüm yayınlama ve dağıtım |
| [Linux değerlendirmesi](Docs/LINUX-DEGERLENDIRME.md) | Linux sürümünün fizibilitesi ve yol haritası |

Program içinde: **Help ▸ Öğretici** (6 bölüm, 21 adım) ve **Help ▸ Kullanım Kılavuzu**.

---

## Kısayollar

| Tuş | |
|---|---|
| `F1` | Tam yakalama → çıktı klasörü + arşiv |
| `F2` | **Bölge ekran görüntüsü** |
| `F3` | Son yakalama yolunu yapıştır |
| `F4` | Kimlik şeritli kare (Atlas) |

Bu dört tuş **global**'dir — UIBUL açıkken her uygulamada çalışır. Başka global kısayol
yoktur; inceleme başlat/durdur, yenile, TXT rapor ve sadece-arşive yakalama üst şeritteki
düğmelerden ve menülerden yapılır.

---

## Geliştirme

```powershell
# Derle ve çalıştır
dotnet run --project UIElementInspector\UIElementInspector\UIElementInspector.csproj

# Setup .exe üret (masaüstüne de kopyalar)
powershell -ExecutionPolicy Bypass -File tools\build-release.ps1
```

### Proje yapısı

```
UIElementInspector/UIElementInspector/   Ana uygulama (WPF)
  Core/Detectors/                        5 algılama motoru
  Core/Models/                           Veri modelleri, ayarlar
  Core/Utils/                            Export, arşiv, ekran görüntüsü, güncelleme
  Services/                              Global kısayol ve fare kancası
  Windows/                               Yardımcı pencereler (öğretici, güncelleme, ayarlar)
Installer/UibulSetup/                    Kurulum programı (WinForms sihirbaz)
tools/build-release.ps1                  Tek komutla setup üretme
Docs/                                    Kullanıcı belgeleri (kurulumla birlikte gider)
```

Öğreticiye yeni bir bölüm eklemek için:
[`Windows/TutorialContent.cs`](UIElementInspector/UIElementInspector/Windows/TutorialContent.cs)
— içerik koddan ayrılmıştır, metin eklemek yeterlidir.

---

© 2026 Emrelic
