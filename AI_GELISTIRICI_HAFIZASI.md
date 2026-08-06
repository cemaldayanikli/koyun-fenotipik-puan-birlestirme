# AI Geliştirici Hafızası — Koyun Fenotipik Puan Birleştirme

Bu dosya, projede yapılan geliştirmelerin kalıcı kaydıdır. Amaç: projeye sonradan bakan
kişinin (veya yapay zekâ asistanının) **neyin neden yapıldığını** koddan okumadan
anlayabilmesi.

## Kurallar

- Her geliştirme, **en yeni en üstte** olacak şekilde bir başlık altında yazılır.
- Her kayıt şunları içerir: **tarih**, **talep/belirti**, **teşhis veya karar**,
  **dokunulan dosyalar**, **doğrulama**, **yayın durumu**.
- Sadece "ne yapıldı" değil, **neden öyle yapıldı** da yazılır.
- Aşağıdaki *Kalıcı referans* bölümü kronolojik değildir, **silinmez**.

---

## Kalıcı referans (kronolojik değil — silme)

| | |
|---|---|
| Teknoloji | Streamlit, tek dosya (`app.py`, ~37 KB) |
| GitHub | `cemaldayanikli/koyun-fenotipik-puan-birlestirme` — **PUBLIC** |
| Yayın | Streamlit Community Cloud, `main` dalını izler → **push = canlıya çıkış** |
| Zincirdeki yeri | **Ortadaki halka** |

### Zincir — iki taraflı bağımlılık

```
Fenotipik_Indeks_App              → Fenotipik_Indeks_YYYY.xlsx
        ↓  (girdi)
BU PROJE                          → Puan_Birlesim_*.xlsx / .csv / .tsv
        ↓  (.tsv ELLE kopyala-yapıştır)
"Damızlık Puanlar" Google Sheet
        ↓
damızlık_reforme_seçim_app        → sahada seçim ekranı (puan + rozet)
```

⚠️ **Çıktı sütun sırasını/başlıklarını değiştirme.** `.tsv` kullanıcı tarafından elle
e-tabloya yapıştırılıyor; biçim değişirse yapıştırma bozulur ve saha ekranı **yanlış
puan gösterir** — hata hiçbir yerde patlamaz, sessizce yanlış olur.
Karşı taraf sayfa adını `PUAN_SAYFA_ADAYLAR` listesiyle eşleştirir
(`Puanlar`, `Puan Birleşim`, `Puanlar 2026`, `Sayfa1`, `Sheet1`).

Biçim değiştirmen gerekiyorsa kardeş projeleri **aynı oturumda** güncelle ve buraya
"hangi tarihten sonraki çıktılar yeni biçimde" diye yaz.

### Bilinen kırılganlıklar

- **Tarih ayrıştırma.** Sayısal yıl (`2017`) Excel seri numarası sanılıp saçma bir
  tarihe dönüşüyordu; `2017-01-01` olarak yorumlanacak şekilde düzeltildi
  (commit `a6205d1`). Doğum tarihi girdilerde **metin, sayı ve gerçek tarih** olarak
  gelebilir — yeni bir ayrıştırma yazarken üçünü de sına.
- **Çoklu yıl ortalaması** dâhil edilen yıl sayısına duyarlıdır. Tek yıl verisi olan
  hayvanın ortalaması o tek yıla eşittir; çok yıllı hayvanlarla aynı kefeye konurken
  bunun farkında ol.

### 🔒 Bu repo public

`.gitignore` veri dosyalarını (`*.xlsx`, `*.xlsm`, `*.csv`, `*.tsv`) kasten dışarıda
tutuyor. **Bu kuralları gevşetme.** (`!requirements.txt` istisnası bilinçlidir.)

---

## 6 Ağustos 2026 — GELİŞTİRME ALTYAPISI KURULDU (kod değişmedi)

**Talep (kullanıcı):** "İlgili klasördeki tüm projelere her proje klasörüne AI
geliştirici hafızası ekle, her projeye GitHub reposu ve commit ekle, ayrıca her projede
*deploy yap* dediğimde otomatik olarak clasp ile push edip deploy yapacak hem de
GitHub'a commit edecek şekilde ayarlama yap."

Bu bir Streamlit projesi olduğu için clasp söz konusu değil; burada "deploy" =
**doğrulama + commit + push**. `app.py` değişmedi.

### Eklenenler

| Dosya | Ne işe yarar |
|---|---|
| `CLAUDE.md` | Proje kuralları: zincirdeki yeri, kırılganlıklar, public repo uyarısı |
| `AI_GELISTIRICI_HAFIZASI.md` | Bu dosya |
| `deploy.ps1` | `py_compile` ile tüm `.py`'yi derler → `git commit` + `git push` |
| `.claude/commands/deploy.md` | "deploy yap" denince izlenecek akış |

### Kararlar ve nedenleri

- **`.gitattributes` eklenmedi.** Apps Script projelerine `* -text` konuldu (clasp'la
  bayt eşitliği için) ama bu depoda `core.autocrlf = true` ve geçmiş LF olarak
  saklanmış; `* -text` eklenseydi bir sonraki commit **tüm dosyaları değişmiş**
  gösterirdi.
- **`.gitignore` ezilmedi, üzerine eklendi.** İlk denemede ortak şablonla
  değiştirilmişti; bu `*.xlsx` / `*.csv` / `*.tsv` gibi **veri dosyası** kurallarını
  siliyordu ve repo public olduğu için hayvan verisi sızabilirdi. Geri alınıp yalnızca
  eksik ortak kurallar sonuna eklendi.
- **`deploy.ps1` önce derler, sonra push eder.** `main`e push = canlıya çıkış olduğu
  için sözdizimi hatası doğrudan sahaya iner.

### Doğrulama (yapıldı)

`deploy.ps1` sözdizimi PowerShell ayrıştırıcısıyla temiz. `git status` yalnızca
beklenen dosyaları gösterdi; veri dosyalarının hiçbiri izlenmeye alınmadı.

**Yayın yapılmadı** — bu oturumda uygulama kodu değişmedi.
