# ŞARTNAME — UIBUL-CLI-0910

## ⓪ KİMLİK — HADDİN (`F13`)

```
SEN        YAPIMCI + MOTOR · UIBUL'un komut satırı çıkışının TEK SAHİBİ
DEĞİLSİN   koordinatör DEĞİLSİN · analizci DEĞİLSİN
ÜSTÜN      UIBUL KOORDINATOR-0910  (local_a11c2f7d-62c5-4e42-b02a-03820968e18b)
ALTIN      kimse
YASAKLARIN iş dağıtmak · başkasının dosyasına yazmak · kapsamı kendi
           başına büyütmek · sürüm numarası artırmak · release yayınlamak
```

## ① NİÇİN VARSIN — ölçülmüş boşluk

`ClaudEmre/ALETLER.md` UIBUL'u **🟡 ADAY** olarak tutuyor. Çantaya girme
şartının ①(ölçülmüş ihtiyaç) ve ③(son doğrulama) ayakları var, **②(çağırma
satırı) YOK.** Ölçüm (2026-09-10, koordinatör):

```
grep -rln "args\[" --include=*.cs UIElementInspector     →  0 isabet
```

Komut satırı arabirimi gerçekten yok. Bugünkü bedeli: UIBUL bir **İNSAN
ALETİ**; her sorgu bir insan turu — *"şunu çalıştır, çıktıyı yapıştır."*

Ölçülmüş talep: EczAsist'te **5+ tkinter penceresi**
(`alis_analiz_gui` · `ana_menu` · `botanik_gui` …) ve klasörde duran
`Inspect.exe` (143 KB, 2026-05-15) — biri masaüstü arayüzünü incelemeye
çalışmış. Tarayıcı araçları tkinter penceresine **ULAŞAMAZ**.

⇒ Bu iş bitince UIBUL **AJAN ALETİ** olur: bütün ClaudEmre projelerinde
oturumlar onu kendi çağırır, insan turu düşer.

## ② İŞİN — sırayla

### 🔴 0a. EN BAŞTA PULL — atlanamaz (Emre'nin emri, 2026-09-10)

```bash
git -C C:/Users/ana/Documents/Projects/Uibul pull --ff-only
git -C C:/Users/ana/Documents/Projects/ClaudEmre pull --ff-only
```

**Niçin — ve bu ÖLÇÜLDÜ, varsayılmadı.** Koordinatör bugün açılışta
ClaudEmre'yi pull etti ve `Already up to date.` aldı. Bir saat sonra
yeniden baktığında depo **5 commit gerideydi** — başka bir makine o arada
push etmişti. Yani *"açılışta pull edildi"* ile *"şu an güncel"* **ayrı
şeylerdir**, ve arada geçen her dakika farkı büyütür.

⚠️ `--ff-only` kasıtlı: yerelde commit'siz iş varsa sessizce merge etmez,
**DURUR.** Durursa kendi başına çözme — koordinatöre bildir (`F16`).
⚠️ `Already up to date.` görmek de bir ölçümdür — **görmeden devam etme.**
🔴 Ve UIBUL deposunu da pull et: koordinatör bu depoya yazıyor
(şartnamen, `CLAUDE.md`, `TESPIH.md` oradan geliyor).

### ⚠️ 0b. SONRA TASARIM NOTU — kod yazmadan ÖNCE onay al

🔴 **Doğrudan koda başlama.** Önce `oturumlar/UIBUL-CLI-0910-ILERLEME.md`
dosyasına **en çok 40 satırlık** bir tasarım notu yaz, commit et ve
koordinatöre bildir. Notta şu dört soru cevaplı olsun:

```
① TEK ÖRNEK KAPISI nasıl aşılıyor    (aşağıdaki 🔴 tuzağa bak)
② çıkış nereye yazılıyor              stdout mu, --out <dosya> mı, ikisi mi
③ JSON şeması ne                      ExportManager'ınkiyle AYNI mı, değilse niçin
④ pencere nasıl adresleniyor          başlık · süreç adı · pid · HWND — hangisi
```

Koordinatör onaylamadan 1'e geçme. **Sebebi:** bu bir yeni işlev ve
tasarımı yanlış seçilirse yazılan kodun tamamı çöpe gider.

### 1. 🔴 TEK ÖRNEK KAPISI — bu işin ASIL ZORLUĞU, ve ÖLÇTÜM

`App.xaml.cs` `OnStartup` içinde:

```csharp
_tekOrnek = new Mutex(true, @"Global\UIElementInspector_TekOrnek", out ilkMi);
if (!ilkMi) { PostMessage(HWND_BROADCAST, …); OncekiniOneGetir(); Shutdown(); return; }
```

⇒ **UIBUL açıkken `UIBUL.exe --json …` çalıştırılırsa ikinci örnek
KENDİNİ KAPATIR ve HİÇBİR ÇIKTI ÜRETMEZ.** Hata da vermez.

Bu, `YASALAR A1`in (sessiz hata) ders kitabı vakasıdır: alet "çalıştı"
görünür, çıktı boş gelir, çağıran oturum *"pencere bulunamadı"* sanır.

**Kabul edilebilir üç yol var; hangisini seçtiğini GEREKÇESİYLE yaz:**

```
(a) CLI kipi mutex'i HİÇ ALMAZ     — komut satırı argümanı varsa
    OnStartup mutex'e girmeden CLI yoluna sapar. En basit; ama iki örnek
    aynı anda UIAutomation kullanır, çakışır mı ÖLÇÜLMELİ.
(b) Çalışan örneğe İSTEK YOLLA     — App.xaml.cs'teki RegisterWindowMessage
    kalıbının aynısı; cevap bir dosyaya/named pipe'a yazılır. En doğru,
    en pahalı.
(c) CLI ayrı bir .exe              — konsol uygulaması, aynı Core/ kodunu
    paylaşır (ikinci proje). Mutex hiç devreye girmez.
```

⚠️ Hangisini seçersen seç, **UIBUL AÇIKKEN ve KAPALIYKEN ayrı ayrı sına**
(`C13`: iki yönde sınanmayan denetim çalışmış sayılmaz).

### 2. Konsol yok — `OutputType=WinExe`

`csproj` `WinExe` üretir; `Console.WriteLine` **hiçbir yere gitmez**.
Standart çıkışa yazacaksan `AttachConsole(ATTACH_PARENT_PROCESS)` gerekir.
⇒ En güvenli teslim: `--out <dosya>` **her zaman** çalışsın; stdout bir
**ek** olsun. Bir oturum dosyayı okuyabilir; boş stdout'u teşhis edemez.

### 3. VAR OLANI KULLAN — yeniden yazma

```
Core/Utils/ExportManager.cs:204   JsonConvert.SerializeObject(elements, Indented, …)
Core/Detectors/IElementDetector   GetAllElements(IntPtr windowHandle, CollectionProfile)
                                  GetElementAtPoint(Point, CollectionProfile)
Newtonsoft.Json 13.0.3            csproj'da ZATEN var
```

⇒ JSON şeması **zaten mevcut**. Yeni bir şema uydurmak, aynı veriyi iki
biçimde saklamak olur ve biri bayatlar. Farklı bir şema gerekiyorsa
**niçin** gerektiğini tasarım notuna yaz.

### 4. Asgari arabirim — bundan fazlasını YAPMA

```
UIBUL.exe --json --pencere "<başlık ya da pid>" [--out <dosya>] [--profil <ad>]
UIBUL.exe --json --liste                        açık pencereleri listele
UIBUL.exe --surum                               sürüm numarası
```

⚠️ Kapsamı büyütme (`G8`). Kısayol, ekran görüntüsü, arşiv **komut
satırına taşınmaz** — bu turda değil.

### 5. Belgeye işle

`README.md`ye kısa bir **Komut satırı** bölümü ve `CLAUDE.md`nin kod
haritasına bir satır. `ClaudEmre/ALETLER.md`ye SEN dokunma — UIBUL'u
🟡 ADAY'dan çantaya taşımak koordinatörün işi.

## ③ YAZMA YETKİSİ

```
🟢 SENİN
   UIElementInspector/UIElementInspector/App.xaml.cs
   UIElementInspector/UIElementInspector/UIElementInspector.csproj
   UIElementInspector/UIElementInspector/Cli/**            (yeni dizin)
   UIElementInspector/UIElementInspector/Core/Utils/ExportManager.cs
   oturumlar/UIBUL-CLI-0910-ILERLEME.md
   README.md · CLAUDE.md  (yalnız komut satırı bölümleri)

🔴 SENİN DEĞİL
   Core/Detectors/**       okuyabilirsin, DEĞİŞTİREMEZSİN (5 motor kararlı)
   Services/** · Windows/** · Installer/**
   oturumlar/TESPIH.md · BEKLEYENLER.md · CEPHANE.md   (koordinatörün)
   oturumlar/UIBUL-ANALIZ-0910*                        (öteki oturumun)
   ClaudEmre/ deposundaki HER ŞEY
```

🔴 Commit **daima pathspec'li**: `git commit -F - -- <yol>`.
`git add -A` yasak — depo tek ağaç, iki oturum aynı anda yazıyor (`F4`).
`add`den önce `git diff -- <yol>` ve tek soru: *"bu satırları BEN mi yazdım?"*

## ④ SENİ BAĞLAYAN YASALAR

```
A1   sessiz hata gürültülüden pahalıdır  ← tek örnek kapısı tam bu sınıf
A3   bir çıktının görünmesi için gereken KAPILARIN HEPSİ sayılmadan
     "iş bitti" denmez
C13  yeni denetim İKİ YÖNDE de sınanmadan "çalışıyor" sayılmaz
     ⇒ UIBUL AÇIKKEN ve KAPALIYKEN
F4   pathspec'li commit; başkasının satırını alma
F5   "yazdım" ile "göründü" ayrı olaylardır — çıktıyı aracın DIŞINDAN ölç
     (başka bir kabuktan oku)
F16  aksaklık raporu BEKLEMEZ — takıldığın an bildir
G8   kapsam isteğine sessizce uyulmaz; büyütme gerekiyorsa İTİRAZ ET
CLAUDE.md ①②③④  bu kod tabanının bilinen dört tuzağı — oku
```

## ⑤ HABERLEŞME

```
🔴 ASIL KANAL DOSYADIR, MESAJ YEDEKTİR (F15)
   Her kalem bitince oturumlar/UIBUL-CLI-0910-ILERLEME.md dosyana YAZ ve
   COMMIT AT (pathspec'li, yalnız kendi dosyan). Mesajı da at ama ONA BEL
   BAĞLAMA.
   Koordinatöre mesaj: mcp__ccd_session_mgmt__send_message
     session_id = local_a11c2f7d-62c5-4e42-b02a-03820968e18b
     (adres bayatlarsa list_sessions ile "UIBUL KOORDINATOR-0910" ARA)
   Ekrana yazdığını koordinatör GÖRMEZ.

🔴 AÇILINCA HEMEN HABER VER — "açıldım, brifingi okudum, şu dosyalar bende".
   Nezaket değil PROTOKOL: koordinatör hangi dosyanın kimde olduğunu
   bilmezse aynı dosyayı ikinci oturuma verir.

🔴 "NE OLDU BİZİM İŞ?" gelirse — iş SÜRÜYOR olsa bile HEMEN üç parçalı
   cevap: "İŞ ÜSTÜNDEYİM · şu aşamadayım · ~şu kadar kaldı".

🔴 AKSAKLIK BEKLEMEZ (F16): başka oturumun dosyası gerekiyor · şartname
   yanlış/eksik · sayı beklenenden ÇOK farklı · iş tahminden ÇOK uzayacak.

🔴🔴 ARIZA ÜÇ YERE BİLDİRİLİR:
   ① ÇÖZ ya da ÇÖZDÜR   ② KOORDİNATÖRE bildir (mesaj gitmiyorsa ilerleme
   dosyası + commit)     ③ KULLANICIYA DA SÖYLE — kendi pencerene açıkça.
   Kanal kırıldığında kullanıcı TEK sağlam alıcıdır.
```

## ⑥ BİTİŞ ÖLÇÜTÜ — SAYIYLA

```
① UIBUL.exe --json --pencere "<X>" başka bir kabuktan koşturuldu ve
   AYRIŞTIRILABİLİR JSON üretti — eleman sayısı > 0, ve o sayı gerçek
   pencereyle uyuşuyor
② aynı komut UIBUL AÇIKKEN de, KAPALIYKEN de koştu   ← iki yön
③ --liste en az 1 pencere döndü
④ dotnet build 0 hata
⑤ hiçbir global kısayol davranışı değişmedi (tek örnek mantığına
   dokunduysan bunu AYRICA sına: F1…F4 hâlâ çalışıyor mu)
```

Teslim raporu SAYIYLA: *"bitirdim"* değil — *"87 eleman döndü, şema
ExportManager'la aynı, UIBUL açıkken (b) yoluyla çalışıyor, F1-F4 sınandı."*
⚠️ Bulamadığını **`bulunamadı`** diye yaz — negatif sonuç da sonuçtur.

## ⑦ DURUM BEYANI — teslimden sonra SUSMA

```
✅ "İŞLERİM BİTTİ — boştayım, yeni iş bekliyorum."
⏳ "BEKLİYORUM: <ne> · <kimden> · <ne zaman tekrar bakacağım>"   ← ÜÇÜ BİRDEN
```

⚠️ Sessizlik bir durum DEĞİLDİR: sustuğunda koordinatör seni "hâlâ
çalışıyor" sayar ve sana iş gelmez (ölçüldü: 52 dakika).

## ⑧ EMEKLİLİK NÖBETİ

```bash
py C:/Users/ana/Documents/Projects/ClaudEmre/kutu/emeklilik.py --nobet --kim "UIBUL-CLI-0910"
```

Arka planda koşar (`Monitor` ile; döngü ÇAĞIRANDA). Sıfır model tokeni.
İŞÇİ mantığı: iş bitti + 45 dk boşta ⇒ EMEKLİ OL · bağlam ≥ 80 K + iş
bitmedi ⇒ İŞ BÖLÜNSÜN. Nöbetçi ötünce **sen karar vermezsin**, durumu
koordinatöre bildirirsin. "ÖLÇÜLEMEDİ" **"devretme" demek değildir.**

## OKUMA LİSTESİ — işe başlamadan

```
CLAUDE.md                                    ← bu projenin künyesi + 4 tuzak
C:/Users/ana/Documents/Projects/ClaudEmre/KISALTMALAR.md
     ⭐ Emre'nin yıldızlı komutları (*mgy · *yyy · *iii · *kii · *ct).
     Emre bunları SANA da yazar; tanımayan oturum komutu ANLAMAZ.
C:/Users/ana/Documents/Projects/ClaudEmre/ALETLER.md   → "🟡 ADAY — UIBUL" bölümü
```
