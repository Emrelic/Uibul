# BEKLEYENLER — UIBUL · Emre'den beklenenler

> koordinatör yazar · her madde başında **kovası** var (`F19`)
> `[izin]` onay · `[karar]` seçim senin · `[eylem]` elini süreceksin ·
> `[cevap]` sorduğumun karşılığı · `[bilgi]` tebligat, bir şey beklenmiyor

---

## B-001 `[cevap]` Projenin amacı — ③ ve ④ yazılı değil

**① ne ölçtüm:** `README.md` amacın iki yüzünü taşıyor — ① ne yapar
(Windows/web arayüzünü içeriden okur + ekran görüntüsü/arşiv) ve ② kim
kullanır (kullanıcı · geliştirici; `Docs/` üç belge). `AMAC.md` yok,
`YOL-HARITASI` yok, `.claude/` altında `.md` yok (arandı, 0 dosya).

**② neyi bulamadım:** ③ *"şu an en çok canını sıkan eksik ne"* ve
④ *"bir şeyden vazgeçmek gerekse neyden vazgeçilir"* — hiçbir dosyada yok.
④ en değerlisi: bütün öncelik kararları ondan türer.

**③ senden ne istiyorum:** iki cümle. Cevap projenin kendi `AMAC.md`sine
yazılacak, ClaudEmre'ye değil.

---

## B-002 `[izin]` Komple analiz yapayım mı? (`E6`)

**① ne ölçtüm:** 87 `.cs` dosyası · 31.281 satır · 5 algılama motoru ·
39 commit. Hiçbiri ClaudEmre gözüyle okunmadı.

**② neyi bulamadım:** analiz yapılmışlığına dair hiçbir iz — ne rapor,
ne `oturumlar/`, ne `BEKLEYENLER` geçmişi.

**③ senden ne istiyorum:** iki şıktan biri —
- **(a)** *"bak"* → eksikler · fazlalar · hatalar · amaca hizmet etmeyen
  yapılar + önerilerim çıkar, **yol haritası** olarak onayına sunulur
- **(b)** *"ben söylerim"* → **kullanıcı-güdümlü kip** (`AMAC.md §5b`):
  sen söylersin, ben önceliklendirir yaparım; iş bitince *"sırada ne var"*
  diye sorarım

---

## B-003 `[karar]` UIBUL'a komut satırı çıkışı yazılsın mı?

**① ne ölçtüm:** `ClaudEmre/ALETLER.md` UIBUL'u **🟡 ADAY** diye tutuyor.
Çantaya girmesini engelleyen tek eksik: *çağırma satırı*. Kendim ölçtüm —
`grep -rln "args\[" --include=*.cs` → **0 isabet**; komut satırı arabirimi
gerçekten yok. Ölçülmüş ihtiyaç yazılı: EczAsist'in 5+ tkinter penceresi
(`alis_analiz_gui` · `ana_menu` · `botanik_gui` …) ve klasöründe duran
`Inspect.exe` — biri masaüstü arayüzünü incelemeye çalışmış, tarayıcı
araçları tkinter penceresine ULAŞAMIYOR.

**② neyi bulamadım:** bu işin daha önce konuşulduğuna dair kayıt yok
(`kutu/SICIL.md`de UIBUL geçmiyor).

**③ senden ne istiyorum:** bu bir **YENİ İŞLEV**, o yüzden sorulur
(`CLAUDEMRE.md §1`). Yazılsın mı? Bugünkü bedeli: her UIBUL sorgusu bir
insan turu ("çalıştır ve yapıştır"). Açılırsa UIBUL insan aletinden
**ajan aletine** döner ve bütün projelerde oturumlar onu kendi çağırır.

---

## B-004 `[karar]` 2026-09-07 gününün hasadı yapılsın mı?

**① ne ölçtüm:** UIBUL bugün sisteme kaydedildi ve tören borcu tablosu
18 iş günü bastı (2025-11-23 … 2026-09-07). 17'si hasat EDİLMEDİ,
gerekçesi yazılı (`ClaudEmre/arsiv/uibul/2026-09-10.md`): o günlerin
hiçbirinde ClaudEmre burada çalışmadı — `oturumlar/`, `AMAC.md`, `arac/`,
`BEKLEYENLER.md`, dördü de yoktu.

**② neyi bulamadım:** 2026-09-07 bir sınır vakası — gerçek bir iş günü ve
ölçülebilir bir teslimi var (`3aec20c`: 12 global kısayol → 4, öğretici +
kılavuz + ayar penceresi + README + TANITIM güncellendi). O gün de
ClaudEmre oturumu değildi, ama commit mesajı ders taşıyor olabilir.

**③ senden ne istiyorum:** hasat edilsin mi, yoksa 17'siyle birlikte
kapatılsın mı? (Kendiliğinden atlamadım — sessiz erteleme yok.)

---

## B-005 `[eylem]` Hazır kıta 0 — ekip gerekiyorsa oturum açman gerekecek

**① ne ölçtüm:** `list_sessions` → 25 oturum. Bu dizinde (`…/Uibul`)
**tek** oturum var: *"Kısayol işlevleri"* (2026-09-07, koşmuyor). Geri
kalan 24'ün hepsi EczAsist ya da ilactarif dizininde — **dizin açılışta
sabitlenir**, o kıtalar bu projeye verilemez.

**② neyi bulamadım:** bu projede hazır kıta havuzu hiç kurulmamış
(`arac/defter.py` yok).

**③ senden ne istiyorum:** **şimdi bir şey yapma.** Bu madde bir
tebligatın eylem yüzü: B-002/B-003 cevaplanınca kaç oturum ve hangi
modelden gerektiğini tek satırla söyleyeceğim. İş tek oturumla yürürse
hiç açılmaz (`G14 ④`: kullanılmayacak yapı kurulmaz).

---

## B-006 `[bilgi]` Başka iki projede hasat borcu var — senin bilgin için

`claudemre-sistemi` 2 gün (09-06, 09-07) · `ilactarif` 2 gün (09-03, 09-04).
İkisi de bu oturumun işi değil; ilgili koordinatöre görünür.
⚠️ Ayrı bir kusur: `TARİH COĞRAFYA SİTESİ` (45 commit) tören borcu
tablosunda **hiç yok** — eklenene kadar borcu ölçülmüyor. Bugün UIBUL'da
aynı kusuru buldum ve düzelttim; atlas'ınki başka makinede (`emrem` profili).
