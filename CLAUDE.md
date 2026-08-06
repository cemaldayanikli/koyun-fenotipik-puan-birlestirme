# Koyun Fenotipik Puan Birleştirme

Birden fazla yıllık fenotipik indeks çıktısını birleştiren, çoklu yıl ortalama puan
hesaplayan ve damızlık/reforme önerisi üreten Streamlit uygulaması.

GitHub: `cemaldayanikli/koyun-fenotipik-puan-birlestirme` (**public**)

## ⚠️ İŞE BAŞLAMADAN ÖNCE

`AI_GELISTIRICI_HAFIZASI.md` dosyasını oku. Geliştirme kararları ve "neden böyle
yapıldı" bilgisi orada, en yeni bölüm en üstte. **İş bitince en üste tarih başlıklı
yeni bir bölüm ekle.**

`README.md` akış şemasını ve kullanıcıya dönük anlatımı tutuyor.

## Zincirdeki yeri — ortadaki halka

```
Fenotipik_Indeks_App                  → Fenotipik_Indeks_YYYY.xlsx   (yılda bir kez)
        ↓
koyun-fenotipik-puan-birleştirme      → Puan_Birlesim_*.xlsx / .csv / .tsv   ◀ BU PROJE
        ↓ (.tsv kopyala-yapıştır)
"Damızlık Puanlar" Google Sheet
        ↓
damızlık_reforme_seçim_app            → sahada seçim ekranı (puan + rozet)
```

İki taraflı bağımlılık — **ikisini de kırmadan değiştir:**

- **Girdi:** `Fenotipik_Indeks_App`'in ürettiği sütun adlarını okur. O proje sütun
  adı değiştirirse burası kırılır.
- **Çıktı:** `.tsv` çıktısı kullanıcı tarafından **elle kopyalanıp** "Damızlık
  Puanlar" e-tablosuna yapıştırılır; oradan `damızlık_reforme_seçim_app` okur.
  O tarafta sayfa adı `PUAN_SAYFA_ADAYLAR` listesiyle (`Puanlar`, `Puan Birleşim`,
  `Puanlar 2026`, `Sayfa1`, `Sheet1`) eşleştirilir. **Sütun sırası ve başlıkları
  değiştirirsen yapıştırma bozulur ve saha ekranı yanlış puan gösterir.**

Çıktı biçimini değiştirmen gerekiyorsa kardeş projeleri de aynı oturumda güncelle
ve hafızaya "hangi tarihten sonraki çıktılar yeni biçimde" diye yaz.

## Proje yapısı

| Dosya | Ne |
|---|---|
| `app.py` | Uygulamanın tamamı, tek dosya (~37 KB) |
| `requirements.txt` | Paket bağımlılıkları |
| `README.md` | Akış şeması + kullanıcı anlatımı |

`app.py` tek dosya; baştan sona okumaya çalışma. `grep -n "^def \|^# ==="` ile harita
çıkarıp ilgili bölümü düzenle.

## Dikkat edilecekler

- **Tarih ayrıştırma kırılgan.** Son commit tam bu yüzden atıldı: sayısal yıl (`2017`)
  Excel seri numarası sanılıp saçma bir tarihe dönüşüyordu, `2017-01-01` olarak
  yorumlanacak şekilde düzeltildi. Girdi dosyalarında doğum tarihi hem metin hem
  sayı hem gerçek tarih olarak gelebilir — yeni bir ayrıştırma yazarken üçünü de sına.
- **Çoklu yıl ortalaması** hangi yılların dâhil edildiğine duyarlıdır. Bir hayvanın
  yalnızca bir yıl verisi varsa ortalaması o tek yıla eşittir — birden çok yılı olan
  hayvanlarla aynı kefeye konurken bunun farkında ol.

## Dağıtım

```powershell
.\deploy.ps1 "Ne değişti"
```

Önce `py_compile` ile derler; hata varsa **hiçbir şey push edilmez**. Sonra
`git commit` + `git push`. Ayrıntı: `.claude/commands/deploy.md`.

### ⚠️ İki nokta

1. **`main`e push = canlıya çıkış.** Streamlit Community Cloud `main` dalını izler ve
   push'tan ~1 dk sonra uygulamayı yeniden başlatır. Yarım işi commit etme.
2. **Bu repo public.** `.gitignore` veri dosyalarını (`*.xlsx`, `*.xlsm`, `*.csv`,
   `*.tsv`) kasten dışarıda tutuyor — hayvan verisi herkese açık olmasın diye.
   **Bu kuralları gevşetme.** (`!requirements.txt` istisnası bilinçlidir.)

## Yerel çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```
