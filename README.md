# Koyun Fenotipik Puan Birleştirme

Birden fazla yıllık fenotipik indeks çıktısını birleştiren, çoklu yıl
ortalama puan hesaplayan, damızlık/reforme önerisi üreten Streamlit uygulaması.

## Akış

```
┌─────────────────────────┐
│ Fenotipik_Indeks_App    │  Her yıl bir kez çalıştırılır.
│ (app.py — ham veri)     │  Çıktı: Fenotipik_Indeks_YYYY.xlsx
└────────────┬────────────┘
             │
   ┌─────────┴───────────┐
   │ Yıllık xlsx'ler     │
   └─────────┬───────────┘
             ▼
┌─────────────────────────┐
│ Bu app                  │  Çoklu yıl birleştirme + ortalama + öneri.
│ (Streamlit)             │  Çıktı: Puan_Birlesim_*.xlsx / .csv / .tsv
└────────────┬────────────┘
             │ (.tsv kopyala-yapıştır)
             ▼
┌─────────────────────────┐
│ "Damızlık Puanlar" Sheet│  Apps Script ID ile bağlanır.
│ (Google Sheets)         │
└────────────┬────────────┘
             ▼
┌─────────────────────────┐
│ Netlify selection app   │  Selection ekranında puan + rozet gösterir.
│ (index.html + kod.gs)   │
└─────────────────────────┘
```

## Lokal çalıştırma

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Community Cloud'a deploy

1. Bu repo'yu GitHub'a push et (public veya private).
2. https://streamlit.io/cloud adresinde "New app" → repo'yu seç → `app.py`'yi göster.
3. Deploy butonuna bas. URL hazır.

## Girdiler

### 1. Yıllık puan dosyaları (zorunlu, çoklu)

`Fenotipik_Indeks_App`'in ürettiği `Fenotipik_Indeks_YYYY*.xlsx` dosyaları.
"Fenotipik Indeks" sayfasında şu sütunları arar:

| Anahtar | Beklenen sütun adı |
|---|---|
| Koyun No | `Koyun No` |
| Puan | `Fenotipik İndeks` |
| Yaş | `Koyun Yaş` (opsiyonel) |
| Doğum Tipi | `Doğum Tipi` (opsiyonel) |

Yıl bilgisi otomatik olarak dosya adından çıkarılır (4 haneli, `20XX`),
arayüzden her dosya için override edilebilir.

### 2. Aktif koyun listesi (opsiyonel)

Çiftleşmesi muhtemel, şu an yaşayan koyunların listesi. Yüklenmezse tüm
puan dosyalarındaki koyunlar gösterilir.

Excel veya CSV. Zorunlu sütun: `Koyun No`.
Opsiyonel sütunlar (varsa otomatik tanınır): `UKN`, `Yaş`, `Doğum Tipi`,
`Padok`, `Irk`.

## Çıktı

### Tablo sütunları

| Sütun | Açıklama |
|---|---|
| Koyun No | İşletme içi numara |
| UKN | Aktif listeden geliyorsa |
| Yaş, Doğum Tipi | Aktif liste önceliklidir, yoksa son yıldan |
| Puan YYYY | Her yıl için ayrı sütun |
| Basit Ort | Eksik yıl = 0 sayılarak ortalama |
| Ağırlıklı Ort | Son yıl daha ağır (sidebar'dan ayarlanır) |
| Yıl Sayısı | Kaç yıl puanı dolu |
| Rank Basit, Rank Ağırlıklı | Yüksekten düşüğe sıra (1=en iyi) |
| Öneri | Damızlık / Reforme / Sınırda |

### İndirme formatları

- **`.xlsx`** — Puan Birleşim + Ayarlar sayfaları
- **`.csv`** — Tek sayfa, Excel uyumlu (UTF-8 BOM)
- **`.tsv`** — Google Sheets "Puanlar" sayfasına yapıştırmak için tab-delimited

## Hesaplamalar

### Basit ortalama
`mean(Puan_yıl1, Puan_yıl2, ...)` — eksik yıllar **0 sayılarak** dahil edilir.
Yani 3 yıllık veride 1 yıl boşsa ortalama düşer.

### Ağırlıklı ortalama
```
Ağ_Ort = Σ(w_i × P_i) / Σ(w_i)
```
Ağırlıklar sidebar'dan ayarlanır:
- `son_yil_sayisi` kadar yıl → `agirlik_son` (default 2.0)
- Sonraki `orta_yil_sayisi` kadar yıl → `agirlik_orta` (default 1.5)
- Geri kalanı → `agirlik_eski` (default 1.0)

Default: 3 yıllık veride sıra 2024:1.0, 2025:1.5, 2026:2.0.

### Öneri kuralları (hibrit)

3 hazır şablon + Özel. Her şablon:

| Parametre | Anlamı |
|---|---|
| `damizlik_min_puan` | Bu puanın altı damızlık değil |
| `damizlik_max_yas` | Bu yaşın üstü damızlık değil |
| `reforme_max_puan` | Bu puanın altı kesin reforme |
| `reforme_min_yas` | Bu yaşın üstü kesin reforme |
| `tek_dogum_cezasi` | Yaş≥X + Doğum Tipi=T + Puan<Y → reforme |

Hesap sırası: Reforme koşulları → Damızlık koşulları → kalan = Sınırda.

## Netlify selection app ile entegrasyon (sıradaki adım)

Sonraki iş paketi:

1. Bu app'in `.tsv` çıktısı yeni bir Google Sheet'in "Puanlar" sayfasına yapıştırılır.
2. Netlify app'in `kod.gs` dosyasına `PUAN_SHEET_ID` sabiti ve `puanGetir()` fonksiyonu eklenir.
3. `index.html` selection kartında her koyunun yıllık puanları, ortalama, rank ve **🟢 Damızlık / 🔴 Reforme / ⚪ Sınırda** rozeti gösterilir.

Bu repo sadece **birleştirme** kısmı.
