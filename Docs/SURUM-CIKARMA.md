# Yeni sürüm çıkarma ve arkadaşlarınıza dağıtma

Bu belge **size** (geliştiriciye) yöneliktir. Kullanıcılar için:
[`TANITIM.md`](TANITIM.md).

---

## Sistem nasıl çalışıyor?

```
  Siz                          GitHub                     Arkadaşınız
  ───                          ──────                     ───────────
  Kod değişikliği
       │
       ▼
  tools\yeni-surum.ps1  ───►  Releases                     UIBUL açılır
  (setup.exe üretir)          (v3.2.0 etiketi              │
       │                       + UIBUL_Setup.exe)          ▼
       └──── yüklersiniz ──────────►                  Releases API'sine bakar
                                                           │
                                                     Yeni sürüm var mı?
                                                           │
                                                    ┌──────┴──────┐
                                                   Yok           Var
                                                    │             │
                                             (hiçbir şey     Pencere açılır
                                              olmaz)         → tek tıkla kurulur
```

Uygulama sürümünü **`UIElementInspector.csproj`** içindeki `<Version>` alanından okur ve
GitHub'daki en son release'in **etiketiyle** karşılaştırır. Etiket sürüm numarası olmak
zorundadır (`v3.2.0` ya da `3.2.0`); release başlığı serbest metindir.

---

## Adım adım: yeni sürüm yayınlamak

### 1. Sürüm numarasını artırın

`UIElementInspector/UIElementInspector/UIElementInspector.csproj`:

```xml
<Version>3.2.0</Version>
<AssemblyVersion>3.2.0.0</AssemblyVersion>
<FileVersion>3.2.0.0</FileVersion>
```

`Installer/UibulSetup/UibulSetup.csproj` içindeki `<Version>` de aynı olmalı
(kurulum penceresinde ve "Uygulamalar" listesinde bu görünür).

### 2. Setup .exe üretin

```powershell
powershell -ExecutionPolicy Bypass -File tools\build-release.ps1
```

Ne yapar:
1. Uygulamayı **self-contained** olarak yayınlar (`dotnet publish`)
2. Belgeleri (`Docs\`) çıktının içine kopyalar
3. Çıktıyı sıkıştırıp installer projesinin içine gömer
4. Installer'ı tek dosyalık `.exe` olarak derler
5. Sonucu `dist\` klasörüne ve **masaüstünüze** kopyalar

Kullanışlı seçenekler:

```powershell
# Masaüstüne kopyalama
powershell -ExecutionPolicy Bypass -File tools\build-release.ps1 -MasaustuneKopyalama

# Sadece installer'ı yeniden derle (uygulama zaten yayınlandıysa — çok daha hızlı)
powershell -ExecutionPolicy Bypass -File tools\build-release.ps1 -SadeceInstaller
```

### 3. GitHub'a release olarak yükleyin

```bash
git add -A
git commit -m "v3.2.0"
git push
gh release create v3.2.0 dist/UIBUL_Setup.exe --title "UIBUL v3.2.0" --notes "Değişiklikler..."
```

`gh` CLI yoksa: GitHub ▸ Releases ▸ *Draft a new release* ▸ etiketi `v3.2.0` yapın,
`UIBUL_Setup.exe` dosyasını sürükleyip bırakın, *Publish release*.

### 4. Bitti

Arkadaşınız programı bir sonraki açtığında güncelleme penceresi kendiliğinden çıkar.

---

## Sürüm notları nasıl yazılmalı?

Release'in **body** alanına yazdığınız metin, güncelleme penceresinde
kullanıcıya **olduğu gibi** gösterilir. Teknik commit mesajı değil, kullanıcının
anlayacağı dilde yazın:

```
Yenilikler
- F12 ile artık tüm pencerenin ekran görüntüsü alınabiliyor
- Ekran görüntüleri artık tarihe göre klasörleniyor

Düzeltmeler
- F9 bazı çoklu monitör kurulumlarında yanlış ekranı yakalıyordu
- Çok uzun element adları pencereyi taşırıyordu
```

---

## İlk paylaşım (arkadaşınız programı hiç kurmadıysa)

İki yol var:

**A. Setup dosyasını doğrudan gönderin** — WhatsApp/Drive/USB fark etmez.
`dist\UIBUL_Setup.exe` tek dosyadır, başka hiçbir şeye gerek yoktur.

**B. Release bağlantısını gönderin** — `https://github.com/Emrelic/Uibul/releases/latest`
Bu daha iyidir: her zaman en güncel sürümü indirirler.

> **Depo public olmalı.** Otomatik güncelleme, GitHub Releases API'sini kimlik
> doğrulaması olmadan sorar. Depo private ise arkadaşınızın token girmesi gerekir
> ve bu akış çalışmaz.

---

## Windows SmartScreen uyarısı

Kurulum dosyası **dijital olarak imzalı değildir** (kod imzalama sertifikası yıllık
ücretlidir). Bu yüzden ilk çalıştırmada Windows şu uyarıyı gösterir:

> *"Windows bilgisayarınızı korudu"*

Arkadaşınıza şunu söyleyin: **Ek bilgi** ▸ **Yine de çalıştır**.

Bu uyarı, dosya birkaç yüz kişi tarafından indirildikçe kendiliğinden kaybolur
(SmartScreen itibar tabanlıdır). İmzalamak isterseniz bir kod imzalama sertifikası
alıp `signtool` ile imzalamanız gerekir — akış aynı kalır, sadece derleme sonrası
bir adım eklenir.

---

## Sorun giderme

| Belirti | Sebep | Çözüm |
|---|---|---|
| Güncelleme penceresi hiç çıkmıyor | Aynı gün zaten bakılmış | `settings.json` → `SonGuncellemeKontrolu` alanını silin, ya da Help ▸ Güncellemeleri kontrol et |
| "Sürüm etiketi anlaşılamadı" | Etiket sürüm numarası değil (`son-surum` gibi) | Etiketi `v3.2.0` biçimine getirin |
| "Bu sürüme kurulum dosyası eklenmemiş" | Release'e `.exe` yüklenmemiş | Release'i düzenleyip `UIBUL_Setup.exe`'yi ekleyin |
| Kurulum "uygulama paketi gömülmemiş" diyor | `build-release.ps1` çalıştırılmadan installer derlenmiş | Betiği baştan çalıştırın |
| Derleme "dosya kilitli" hatası veriyor | UIBUL çalışıyor | Programı kapatın, tekrar deneyin |

---

## Dosya haritası

| Yol | Ne |
|---|---|
| `UIElementInspector/UIElementInspector/` | Ana uygulama (WPF) |
| `UIElementInspector/.../Core/Utils/UpdateService.cs` | Güncelleme kontrolü ve indirme |
| `UIElementInspector/.../Windows/UpdateWindow.xaml` | Güncelleme penceresi |
| `UIElementInspector/.../Windows/TutorialWindow.xaml` | Öğretici penceresi (motor) |
| `UIElementInspector/.../Windows/TutorialContent.cs` | **Öğretici metinleri** — yeni özellik eklerken buraya adım ekleyin |
| `Installer/UibulSetup/` | Kurulum programı (WinForms sihirbaz) |
| `Installer/UibulSetup/Gereklilikler.cs` | Sistem denetimleri |
| `Installer/UibulSetup/Kurulum.cs` | Dosya yerleştirme, kısayol, kayıt defteri |
| `tools/build-release.ps1` | Tek komutla setup üretme |
| `Docs/` | Kullanıcı belgeleri (kurulumla birlikte gider) |
