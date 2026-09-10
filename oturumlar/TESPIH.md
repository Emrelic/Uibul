# TESPİH — UIBUL

> kurulum: 2026-09-10 · `UIBUL KOORDINATOR-0910`
> Sıralama ölçütü `FAYDA ÷ EMEK` — `HEDEF − KEŞKİNLİK` değil.
> ⚠️ Amacın ③④ yüzü (`BEKLEYENLER.md` B-001) hâlâ cevapsız; cevap gelince
> 3. ve 4. sıralar yeniden ölçülür.

| # | iş | keskinlik | hedef | fayda / niçin bu sırada | durum |
|---|-----|-----------|-------|------------------------|-------|
| 1 | UIBUL'a komut satırı çıkışı (`--json <pencere>`) — `UIBUL-CLI-0910` | %0 | %90 | 🔓 **EN BÜYÜK DARBOĞAZ**: `ALETLER.md`de UIBUL 🟡 ADAY, çantaya girmesini engelleyen TEK eksik bu. Açılırsa UIBUL insan aletinden **ajan aletine** döner ve bütün ClaudEmre projeleri onu kendi çağırır. Emre onayladı (*"yazılsın"*) | ⏳ oturum bekliyor |
| 2 | Komple analiz — `UIBUL-ANALIZ-0910` | %0 | %80 | 87 `.cs` / 31.281 satır hiç ClaudEmre gözüyle okunmadı. Emre onayladı (*"bak"*). Çıktısı **yol haritası** olarak Emre'nin onayına sunulacak | ⏳ oturum bekliyor |
| 3 | `AMAC.md` — ③ en çok canını sıkan eksik, ④ öncelik ölçütü | %50 | %95 | 🔓 **DARBOĞAZ**: 4. sıradan sonrasının sırası buradan türer. README ①(ne yapar) ve ②(kim kullanır) yüzlerini taşıyor; ③④ yazılı değil ⇒ ekip yanlış yöne verimli çalışır | 🟡 Emre'de (B-001) |
| 4 | Analiz raporu → **YOL HARİTASI** → Emre onayı | %0 | %90 | Onay gelmeden 2. dalga ekip kurulmaz (`AMAC.md §5`) | ⚪ 2'yi bekliyor |
| 5 | UIBUL'u `ALETLER.md`de 🟡 ADAY'dan **çantaya** taşı | %0 | %70 | 1. sıra bitince tek satırlık iş; ama üç şartın üçü de doldurulacak: ölçülmüş ihtiyaç · **çağırma satırı** · son doğrulama TARİHLE | ⚪ 1'i bekliyor |

## ✅ Bugün kapanan kalemler

- **UIBUL sisteme kaydedildi** (`ClaudEmre@aaa6181`) — `PROJELER.md` ·
  `proje_takma.json` · `toren_borcu.py` tablosu · `arsiv/uibul/`.
  🔴 Bulunan kusur: tabloda olmayan projenin hasat borcu **hiç ölçülmüyordu**;
  2025-11'den beri süren bir depo sisteme görünmezdi.
- **18 günlük hasat borcu kapatıldı** — 17'si gerekçesiyle
  (`arsiv/uibul/2026-09-10.md`), 1'i (2026-09-07) Emre'nin kararıyla
  **hasat edildi**: `ClaudEmre@3131d5d`, iki evrensel ders, biri terfi adayı.
- **Koordinatörlük kuruldu** — `CLAUDE.md` · `TESPIH.md` · `BEKLEYENLER.md` ·
  `CEPHANE.md` + iki şartname.

## Bugünkü ölçüm — bu tablonun dayanağı

```
git log --oneline | wc -l              39 commit · ilk 2025-11-23
git log -1 --date=short                2026-09-07  3aec20c
find … -name "*.cs" | wc -l            87 dosya · 31.281 satır
grep -rln "args\[" --include=*.cs      0 isabet  ⇒ komut satırı YOK
git status --short                     temiz
list_sessions (cwd=Uibul)              1 oturum, koşmuyor · hazır kıta 0
kutu/ozet.py "uibul"                   0 paket · 0 işlenmemiş · 0 cevapsız sohbet
denetle_sistem.py                      temiz ✓ (9 kontrol)
```

## Ekip — dosya alanları KESİŞMİYOR

| oturum | rol | model | dosyaları | niçin ayrı |
|---|---|---|---|---|
| `UIBUL KOORDINATOR-0910` | koordinatör | Opus | `TESPIH` · `BEKLEYENLER` · `CEPHANE` · şartnameler | — |
| `UIBUL-CLI-0910` | yapımcı/motor | Opus | `App.xaml.cs` · `csproj` · `Cli/**` · `ExportManager.cs` | tek örnek mutex'i + WinExe konsolu: hata maliyeti yüksek, kusur **sessiz** olur (`M1`) |
| `UIBUL-ANALIZ-0910` | araştırıcı/denetçi | Opus | `ANALIZ-RAPORU-0910.md` **yalnız** — koda DOKUNMAZ | kendi çıktısını sorgulaması gereken iş (`M3`) |

⚠️ Tek kesişme riski: analizci `App.xaml.cs`/`ExportManager.cs` hakkında
bulgu yazarken CLI oturumu onları değiştiriyor olabilir. Çare şartnameye
kondu — analizci **ölçtüğü commit sha'yı** raporun başına yazar (`B3`).
