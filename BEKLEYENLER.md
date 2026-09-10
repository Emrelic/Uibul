# BEKLEYENLER — UIBUL · Emre'den beklenenler

> koordinatör yazar · her madde başında **kovası** var (`F19`)
> `[izin]` onay · `[karar]` seçim senin · `[eylem]` elini süreceksin ·
> `[cevap]` sorduğumun karşılığı · `[bilgi]` tebligat, bir şey beklenmiyor

---

## 🔴 AÇIK

### B-001 `[cevap]` Projenin amacı — ③ ve ④ yazılı değil

**① ne ölçtüm:** `README.md` amacın iki yüzünü taşıyor — ① ne yapar
(Windows/web arayüzünü içeriden okur + ekran görüntüsü/arşiv) ve ② kim
kullanır (kullanıcı · geliştirici; `Docs/` üç belge). `AMAC.md` yok,
`YOL-HARITASI` yok, `.claude/` altında `.md` yok (arandı, 0 dosya).

**② neyi bulamadım:** ③ *"şu an en çok canını sıkan eksik ne"* ve
④ *"bir şeyden vazgeçmek gerekse neyden vazgeçilir"* — hiçbir dosyada yok.
④ en değerlisi: bütün öncelik kararları ondan türer, ve analiz oturumunun
*"amaca hizmet etmeyenler"* başlığı da ondan hüküm alacak.

**③ senden ne istiyorum:** iki cümle. Cevap projenin kendi `AMAC.md`sine
yazılacak, ClaudEmre'ye değil (`AMAC.md §4`: ClaudEmre yöntem taşır,
proje içeriği taşımaz).

---

### B-005 `[eylem]` İKİ OTURUM AÇ — panolar aşağıda, tek kopyala

**① ne ölçtüm:** `list_sessions` → 25 oturum. Bu dizinde (`…/Uibul`)
**tek** oturum var: *"Kısayol işlevleri"* (2026-09-07, koşmuyor). Geri
kalan 24'ün hepsi EczAsist ya da ilactarif dizininde — **dizin açılışta
sabitlenir**, o kıtalar bu projeye verilemez. Hazır kıta: **0**.

**② neyi bulamadım:** bu projede hazır kıta havuzu hiç kurulmamış
(`arac/defter.py` yok). Sistem oturum AÇAMAZ, model SEÇEMEZ — o iş senin.

**③ senden ne istiyorum:** iki yeni oturum aç. Her biri için **tek
kopyalama** yeter; sen yalnız MODEL ve DİZİNİ ayarlıyorsun, ikisi de
bloğun ilk iki satırında yazılı.

#### PANO 1 — komut satırı çıkışı

```
MODEL   Opus  — tek örnek mutex'i + WinExe konsolu; yanlış tasarım SESSİZ
                çalışmayan bir alet üretir, hata maliyeti yüksek (M1)
DİZİN   C:\Users\ana\Documents\Projects\Uibul

Adın UIBUL-CLI-0910. Koordinatör oturum UIBUL KOORDINATOR-0910.

🔴 İLK İŞ — HİÇBİR ŞEYE BAŞLAMADAN PULL:
   git -C C:/Users/ana/Documents/Projects/Uibul pull --ff-only
   git -C C:/Users/ana/Documents/Projects/ClaudEmre pull --ff-only
   "Already up to date." görmek de bir ölçümdür — GÖRMEDEN devam etme.
   --ff-only DURURSA kendin çözme, koordinatöre bildir.

Sonra CLAUDE.md'yi baştan sona oku, sonra oturumlar/UIBUL-CLI-0910.md —
şartnamen odur, ona göre çalış.
İlk işin: koordinatöre açılış mesajı at (mcp__ccd_session_mgmt__send_message,
session_id = local_a11c2f7d-62c5-4e42-b02a-03820968e18b):
   "açıldım · brifingi okudum · dosyalarım: App.xaml.cs, csproj, Cli/**,
    Core/Utils/ExportManager.cs · oturum kimliğim: <ÖLÇ>"
🔴 session_id'ini mcp__ccd_session_mgmt__get_session("self") ile ÖLÇ —
   scratchpad yolundaki UUID DEĞİLDİR.
🔴 Kod yazmadan ÖNCE tasarım notu ver ve onay bekle (şartname ② madde 0).
```

#### PANO 2 — komple analiz

```
MODEL   Opus  — kendi çıktısını sorgulaması gereken iş; küçük modele
                verilmez (M3). Güç etiketi ve negatif sonuç hükmü gerekiyor
DİZİN   C:\Users\ana\Documents\Projects\Uibul

Adın UIBUL-ANALIZ-0910. Koordinatör oturum UIBUL KOORDINATOR-0910.

🔴 İLK İŞ — HİÇBİR ŞEYE BAŞLAMADAN PULL:
   git -C C:/Users/ana/Documents/Projects/Uibul pull --ff-only
   git -C C:/Users/ana/Documents/Projects/ClaudEmre pull --ff-only
   Bayat ağacı analiz edersen raporun DOĞDUĞU AN bayat olur ve hiçbir
   alarm ötmez. Raporunun başındaki sha PULL SONRASI olacak.

Sonra CLAUDE.md'yi baştan sona oku, sonra oturumlar/UIBUL-ANALIZ-0910.md —
şartnamen odur, ona göre çalış.
İlk işin: koordinatöre açılış mesajı at (mcp__ccd_session_mgmt__send_message,
session_id = local_a11c2f7d-62c5-4e42-b02a-03820968e18b):
   "açıldım · brifingi okudum · dosyam: oturumlar/ANALIZ-RAPORU-0910.md
    (koda DOKUNMUYORUM) · ölçtüğüm commit: <git rev-parse --short HEAD>
    · oturum kimliğim: <ÖLÇ>"
🔴 session_id'ini mcp__ccd_session_mgmt__get_session("self") ile ÖLÇ —
   scratchpad yolundaki UUID DEĞİLDİR.
🔴 HİÇBİR KOD DOSYASINA YAZMA — sen ölçer ve rapor edersin.
```

---

## ✅ KAPANDI

### B-002 `[izin]` Komple analiz yapayım mı? → **"Bak — analiz yap"** (2026-09-10)
Analiz `UIBUL-ANALIZ-0910` oturumuna verildi; şartname yazıldı. Çıktısı
**yol haritası** olarak onayına sunulacak — onay gelmeden 2. dalga ekip
kurulmaz.

### B-003 `[karar]` UIBUL'a komut satırı çıkışı yazılsın mı? → **"Yazılsın"** (2026-09-10)
`UIBUL-CLI-0910` oturumuna verildi. Kapsam bilerek dar tutuldu (`G8`):
`--json --pencere` · `--json --liste` · `--surum`. Kısayol, ekran
görüntüsü ve arşiv komut satırına **taşınmıyor** — bu turda değil.

🔴 **Ve bu iş açılırken bir tuzak ÖLÇÜLDÜ, şartnameye kondu:**
`App.xaml.cs` `Global\UIElementInspector_TekOrnek` mutex'ini alıyor;
alamayan örnek `Shutdown()` çağırıyor. Yani UIBUL **açıkken**
`UIBUL.exe --json …` çalıştırılırsa ikinci örnek kendini kapatır ve
**hiçbir çıktı üretmeden, hata da vermeden** ölür. Üç çözüm yolu
şartnamede duruyor; oturum hangisini seçtiğini gerekçesiyle yazacak.

### B-004 `[karar]` 2026-09-07 gününün hasadı → **"Hasat et"** (2026-09-10)
Yapıldı: `ClaudEmre@3131d5d`. `3aec20c` okundu, **iki evrensel ders**
çıktı:
- *Yeni bir varsayılan, diske yazılmış eski değeri değiştirmez* —
  🎓 **terfi adayı** (aynı şekil ClaudEmre'de 2 Eylül'de ısırdı ⇒ iki
  AYRI proje, çekirdeğe giriş şartı karşılanıyor)
- *Global bir kaynağı rezerve etmenin bedelini, kullanmayan herkes öder*

Karar doğru çıktı: tek commit'ten iki ders, biri terfi adayı.

---

## ℹ️ BİLGİ

### B-006 `[bilgi]` Başka iki projede hasat borcu var

`claudemre-sistemi` 2 gün (09-06, 09-07) · `ilactarif` 2 gün (09-03, 09-04).
İkisi de bu oturumun işi değil; ilgili koordinatöre görünür.

⚠️ Ayrı bir kusur, **aynı sınıftan**: `TARİH COĞRAFYA SİTESİ` (45 commit)
tören borcu tablosunda **hiç yok** — eklenene kadar borcu ölçülmüyor.
Bugün UIBUL'da tam aynı kusuru buldum ve düzelttim; atlas'ınki başka
makinede (`emrem` profili), oradaki koordinatörün işi.

### B-007 `[bilgi]` Cephane kaydedildi
20x Max · limit rahat ⇒ **tam düzen** (3 eşzamanlı). `oturumlar/CEPHANE.md`.
Limit tazelenince ya da daralırsa haber ver, vitesi değiştiririm.
