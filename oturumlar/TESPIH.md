# TESPİH — UIBUL

> kurulum: 2026-09-10 · koordinatör oturumu · **GEÇİCİ**: amaç soruları
> (③ en çok canını sıkan eksik · ④ öncelik ölçütü) cevaplanınca yeniden sıralanır.
> Sıralama ölçütü `FAYDA ÷ EMEK` — `HEDEF − KEŞKİNLİK` değil.

| # | iş | keskinlik | hedef | fayda / niçin bu sırada | durum |
|---|---|---|---|---|---|
| 1 | `AMAC.md` — ③ en çok canını sıkan eksik, ④ öncelik ölçütü | %50 | %95 | 🔓 **DARBOĞAZ**: aşağıdaki her satırın sırası buradan türer. README ①(ne yapar) ve ②(kim kullanır) yüzlerini zaten taşıyor; ③④ yazılı değil ⇒ ekip yanlış yöne verimli çalışır | 🔵 sırada |
| 2 | UIBUL'a komut satırı çıkışı (`--json <pencere>`) | %0 | %90 | 🔓 **EN BÜYÜK DARBOĞAZ**: `ALETLER.md`de UIBUL 🟡 ADAY olarak duruyor ve çantaya girmesini engelleyen TEK eksik bu. Bugün her sorgu bir insan turu ("çalıştır, yapıştır"). Açılırsa UIBUL **ajan aleti** olur ve EczAsist'in 5+ tkinter penceresi dâhil bütün ClaudEmre projeleri onu kendi çağırır. ⚠️ YENİ İŞLEV ⇒ sorulmadan başlanmaz (`CLAUDEMRE.md §1`) | 🔵 sırada |
| 3 | Analiz raporu (eksik · fazla · hata · amaca hizmet etmeyen) | %0 | %80 | 87 `.cs` / 31.281 satır hiç ClaudEmre gözüyle okunmadı. ⚠️ İZİNLİ (`E6`) — kullanıcı istemezse ATLANIR ve kullanıcı-güdümlü kipe geçilir | ⚪ izin bekliyor |
| 4 | ClaudEmre altyapısı (`arac/defter.py` · `BEKLEYENLER.md` · oturum şartnameleri) | %0 | %80 | Yalnız **ekip kurulacaksa** gerekir. Tek oturumla yürüyecekse kurulmaz — `G14 ④`: kullanılmayacak yapı kurulmaz | ⚪ bekletildi |
| 5 | 2026-09-07 gününün hasadı | %0 | %60 | Gerçek bir iş günü, ölçülebilir teslimi var (12 global kısayol → 4). Ama ClaudEmre oturumu değildi ⇒ kendiliğinden atlanmadı, **soruldu** | ⚪ karar bekliyor |

## Bugünkü ölçüm — bu tablonun dayanağı

```
git log --oneline | wc -l              39 commit · ilk 2025-11-23
git log -1 --date=short                2026-09-07  3aec20c
find … -name "*.cs" | wc -l            87 dosya · 31.281 satır
grep -rln "args\[" --include=*.cs      0 isabet  ⇒ komut satırı YOK (kendim ölçtüm)
git status --short                     temiz
list_sessions (cwd=Uibul)              1 oturum, koşmuyor · hazır kıta 0
kutu/ozet.py "uibul"                   0 paket · 0 işlenmemiş · 0 cevapsız sohbet
```

## Kapanan kalemler

- ✅ **UIBUL sisteme kaydedildi** (2026-09-10, `aaa6181`) — `PROJELER.md` ·
  `proje_takma.json` · `toren_borcu.py` tablosu · `arsiv/uibul/`.
  18 günlük geriye dönük hasat borcu **gerekçesiyle** kapatıldı:
  `ClaudEmre/arsiv/uibul/2026-09-10.md`
