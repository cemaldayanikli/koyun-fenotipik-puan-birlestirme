"""
Koyun Fenotipik Puan Birleştirme (Streamlit)
=============================================
Birden fazla yıllık Fenotipik İndeks çıktısını (Fenotipik_Indeks_Indeks_App'in
ürettiği .xlsx dosyaları) birleştirir, çoklu yıl ortalama puanı hesaplar,
şu an yaşayan koyun listesine göre filtreler, damızlık/reforme önerisi üretir.

Çıktı: makrolu Excel'in "Puan" sayfası gibi tablo + "Puanlar Sheet"ine
yapıştırma için CSV.

Çalıştırma:
    streamlit run app.py
"""

import io
import re
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st


# ═══════════════════════════════════════════════════════════════════════════
# SABİT TANIMLAR
# ═══════════════════════════════════════════════════════════════════════════

# Fenotipik_Indeks_App'in çıktısındaki temel sütunlar
EXPECTED_COLS = {
    'koyun_no':   ['Koyun No', 'koyun_no', 'koyun no'],
    'koyun_yas':  ['Koyun Yaş', 'Koyun Yas', 'koyun_yas', 'yaş', 'yas'],
    'dogum_tipi': ['Doğum Tipi', 'Dogum Tipi', 'dogum_tipi'],
    'puan':       ['Fenotipik İndeks', 'Fenotipik Indeks', 'Puan', 'İndeks'],
}

# Önceden tanımlı öneri şablonları
SABLONLAR = {
    'Standart': {
        'damizlik_min_puan':    60.0,
        'damizlik_max_yas':     6,
        'reforme_max_puan':     40.0,
        'reforme_min_yas':      7,
        'tek_dogum_cezasi':     True,   # Yaş≥5 + Doğum Tipi=T + Puan<50 → reforme
        'tek_dogum_min_yas':    5,
        'tek_dogum_max_puan':   50.0,
    },
    'Sıkı': {
        'damizlik_min_puan':    70.0,
        'damizlik_max_yas':     5,
        'reforme_max_puan':     50.0,
        'reforme_min_yas':      6,
        'tek_dogum_cezasi':     True,
        'tek_dogum_min_yas':    4,
        'tek_dogum_max_puan':   55.0,
    },
    'Gevşek': {
        'damizlik_min_puan':    45.0,
        'damizlik_max_yas':     7,
        'reforme_max_puan':     30.0,
        'reforme_min_yas':      8,
        'tek_dogum_cezasi':     False,
        'tek_dogum_min_yas':    5,
        'tek_dogum_max_puan':   50.0,
    },
}


# ═══════════════════════════════════════════════════════════════════════════
# YARDIMCI FONKSİYONLAR
# ═══════════════════════════════════════════════════════════════════════════

def yil_bul_dosya_adindan(name: str) -> int | None:
    """Dosya adından 4 haneli yıl çıkar (2020-2099)."""
    m = re.search(r'(20\d{2})', name)
    return int(m.group(1)) if m else None


def sutun_eslestir(df: pd.DataFrame) -> dict:
    """DataFrame sütunlarını standart anahtarlara eşle."""
    cm = {}
    for key, candidates in EXPECTED_COLS.items():
        for col in df.columns:
            col_clean = str(col).strip()
            if col_clean in candidates:
                cm[key] = col
                break
            # case-insensitive fallback
            col_lower = col_clean.lower()
            for c in candidates:
                if c.lower() == col_lower:
                    cm[key] = col
                    break
            if key in cm:
                break
    return cm


def puan_dosyasi_oku(file, yil_override: int | None = None) -> tuple[pd.DataFrame, int]:
    """Tek bir Fenotipik_Indeks_*.xlsx dosyasını oku, gerekli sütunları çek."""
    try:
        df = pd.read_excel(file, sheet_name='Fenotipik Indeks', engine='openpyxl')
    except Exception:
        file.seek(0)
        df = pd.read_excel(file, engine='openpyxl')

    df.columns = [str(c).strip() for c in df.columns]
    cm = sutun_eslestir(df)

    eksik = [k for k in ('koyun_no', 'puan') if k not in cm]
    if eksik:
        raise ValueError(f"Eksik sütunlar: {eksik}. Bulunanlar: {list(df.columns)[:8]}...")

    out = pd.DataFrame({
        'Koyun No':   pd.to_numeric(df[cm['koyun_no']], errors='coerce').astype('Int64'),
        'Puan':       pd.to_numeric(df[cm['puan']], errors='coerce'),
        'Koyun Yaş':  pd.to_numeric(df[cm['koyun_yas']], errors='coerce') if 'koyun_yas' in cm else pd.NA,
        'Doğum Tipi': df[cm['dogum_tipi']].astype(str).str.strip() if 'dogum_tipi' in cm else '',
    })
    out = out.dropna(subset=['Koyun No']).reset_index(drop=True)

    yil = yil_override if yil_override else yil_bul_dosya_adindan(getattr(file, 'name', ''))
    if not yil:
        yil = datetime.now().year
    return out, yil


def aktif_liste_oku(file) -> pd.DataFrame:
    """Aktif koyun listesi dosyasını oku — Koyun No içeren basit tablo."""
    name = getattr(file, 'name', '').lower()
    if name.endswith('.csv'):
        try:
            df = pd.read_csv(file, encoding='utf-8-sig')
        except UnicodeDecodeError:
            file.seek(0)
            df = pd.read_csv(file, encoding='cp1254')
    else:
        df = pd.read_excel(file, engine='openpyxl')

    df.columns = [str(c).strip() for c in df.columns]
    # Koyun No bulmaya çalış
    koyun_kol = None
    for col in df.columns:
        cl = col.lower().replace('ı', 'i')
        if cl in ('koyun no', 'koyun_no', 'koyunno', 'no', 'kn'):
            koyun_kol = col
            break
    if koyun_kol is None:
        raise ValueError(f"'Koyun No' sütunu bulunamadı. Sütunlar: {list(df.columns)}")

    out = pd.DataFrame({
        'Koyun No': pd.to_numeric(df[koyun_kol], errors='coerce').astype('Int64'),
    })
    # Diğer sütunlar varsa al (Yaş, Doğum Tipi, UKN, Padok vs.)
    for col in df.columns:
        if col == koyun_kol:
            continue
        cl = col.lower().replace('ı', 'i').replace('ğ', 'g')
        if 'ukn' in cl or 'ulusal' in cl:
            out['UKN'] = df[col].astype(str).str.strip()
        elif 'yas' in cl or 'yaş' in cl or 'cag' in cl or 'çağ' in cl:
            out['Aktif Yaş'] = pd.to_numeric(df[col], errors='coerce')
        elif 'dogum tipi' in cl or 'doğum tipi' in cl:
            out['Aktif Doğum Tipi'] = df[col].astype(str).str.strip()
        elif 'padok' in cl:
            out['Padok'] = df[col].astype(str).str.strip()
        elif 'irk' in cl or 'cins' in cl:
            out['Irk'] = df[col].astype(str).str.strip()
    out = out.dropna(subset=['Koyun No']).drop_duplicates('Koyun No').reset_index(drop=True)
    return out


def birlestir(puan_dosyalari: dict, aktif: pd.DataFrame | None) -> pd.DataFrame:
    """puan_dosyalari: {yil: DataFrame}. Aktif liste varsa onunla filtrele."""
    if not puan_dosyalari:
        return pd.DataFrame()

    yillar = sorted(puan_dosyalari.keys())
    # Tüm koyunlardan oluşan iskelet
    if aktif is not None and len(aktif) > 0:
        merged = aktif.copy()
    else:
        all_kn = set()
        for yil in yillar:
            all_kn.update(puan_dosyalari[yil]['Koyun No'].dropna().tolist())
        merged = pd.DataFrame({'Koyun No': sorted(all_kn)})

    # Her yıl için puan sütunu ekle
    for yil in yillar:
        df = puan_dosyalari[yil][['Koyun No', 'Puan']].rename(columns={'Puan': f'Puan {yil}'})
        merged = merged.merge(df, on='Koyun No', how='left')

    # En güncel yıldan yaş + doğum tipi al (en yeni yıl)
    son_yil = yillar[-1]
    son_df = puan_dosyalari[son_yil][['Koyun No', 'Koyun Yaş', 'Doğum Tipi']]
    merged = merged.merge(son_df, on='Koyun No', how='left', suffixes=('', '_son'))

    # Aktif liste yaşı varsa onu öncele
    if 'Aktif Yaş' in merged.columns:
        merged['Yaş'] = merged['Aktif Yaş'].fillna(merged.get('Koyun Yaş'))
    else:
        merged['Yaş'] = merged.get('Koyun Yaş')

    if 'Aktif Doğum Tipi' in merged.columns:
        merged['Doğum Tipi'] = merged['Aktif Doğum Tipi'].replace('', np.nan).fillna(merged.get('Doğum Tipi', ''))
    elif 'Doğum Tipi' not in merged.columns:
        merged['Doğum Tipi'] = ''

    return merged, yillar


def ortalamalari_hesapla(df: pd.DataFrame, yillar: list[int], agirlik_son: float = 2.0,
                          agirlik_orta: float = 1.5, agirlik_eski: float = 1.0,
                          son_yil_sayisi: int = 1, orta_yil_sayisi: int = 1) -> pd.DataFrame:
    """İki ortalama sistemi:
    - Basit ortalama: eksik yıl = 0 (ortalamayı düşürür)
    - Ağırlıklı ortalama: son yıllar daha ağır, NaN olanlar 0 sayılır.
    """
    if not yillar:
        df['Basit Ort'] = np.nan
        df['Ağırlıklı Ort'] = np.nan
        return df

    puan_kols = [f'Puan {y}' for y in yillar]
    puan_matrix = df[puan_kols].apply(pd.to_numeric, errors='coerce')

    # Basit ortalama — eksik yıl 0 olarak ortalamaya dahil
    df['Basit Ort'] = puan_matrix.fillna(0).mean(axis=1).round(2)

    # Ağırlıklı ortalama
    n = len(yillar)
    sirali = list(range(n))  # 0=en eski, n-1=en yeni
    agirliklar = []
    for i in sirali:
        # En yeni `son_yil_sayisi` kadar yıl → agirlik_son
        # Sonraki `orta_yil_sayisi` kadar yıl → agirlik_orta
        # Geri kalanı → agirlik_eski
        konum_son = n - 1 - i
        if konum_son < son_yil_sayisi:
            agirliklar.append(agirlik_son)
        elif konum_son < son_yil_sayisi + orta_yil_sayisi:
            agirliklar.append(agirlik_orta)
        else:
            agirliklar.append(agirlik_eski)
    toplam_agirlik = sum(agirliklar)
    if toplam_agirlik > 0:
        weighted = puan_matrix.fillna(0).mul(agirliklar, axis=1).sum(axis=1) / toplam_agirlik
        df['Ağırlıklı Ort'] = weighted.round(2)
    else:
        df['Ağırlıklı Ort'] = 0.0

    # Doluluk (kaç yıl puanı var)
    df['Yıl Sayısı'] = puan_matrix.notna().sum(axis=1).astype(int)

    # Rank — Basit ortalamaya göre (yüksekten düşüğe, 1=en iyi)
    df['Rank Basit'] = df['Basit Ort'].rank(method='min', ascending=False).astype('Int64')
    df['Rank Ağırlıklı'] = df['Ağırlıklı Ort'].rank(method='min', ascending=False).astype('Int64')

    return df, agirliklar


def oneri_uret(df: pd.DataFrame, sablon: dict, hangi_ort: str = 'Basit Ort') -> pd.Series:
    """Hibrit öneri rozeti: 'Damızlık' / 'Reforme' / 'Sınırda'."""
    p = df[hangi_ort].fillna(0)
    y = df['Yaş'].fillna(0)
    dt = df['Doğum Tipi'].astype(str).str.strip().str.upper()

    onerii = pd.Series('Sınırda', index=df.index, dtype=object)

    # Damızlık koşulu
    damizlik_mask = (p >= sablon['damizlik_min_puan']) & (y <= sablon['damizlik_max_yas']) & (y > 0)

    # Reforme koşulu — birkaç alternatif
    reforme_mask = (p <= sablon['reforme_max_puan']) | (y >= sablon['reforme_min_yas'])
    if sablon.get('tek_dogum_cezasi'):
        tek_dogum_reforme = (
            (y >= sablon['tek_dogum_min_yas']) &
            (dt == 'T') &
            (p < sablon['tek_dogum_max_puan'])
        )
        reforme_mask = reforme_mask | tek_dogum_reforme

    onerii.loc[reforme_mask] = 'Reforme'
    onerii.loc[damizlik_mask & ~reforme_mask] = 'Damızlık'
    # Damızlık ve reforme aynı anda işaretliyse → reforme öncelikli (zaten üst sırada)
    return onerii


def excel_olarak_indir(df: pd.DataFrame, yillar: list[int], agirliklar: list[float],
                       sablon: dict, sablon_adi: str) -> bytes:
    """Sonuç tablosunu Excel olarak paketle."""
    bio = io.BytesIO()
    with pd.ExcelWriter(bio, engine='xlsxwriter') as xw:
        df.to_excel(xw, sheet_name='Puan Birleşim', index=False)

        # Sütun genişlikleri
        ws = xw.sheets['Puan Birleşim']
        for i, col in enumerate(df.columns):
            data_max = max((len(str(v)) for v in df[col].astype(str).tolist()), default=10) if len(df) else 10
            header_len = len(str(col)) + 2
            width = max(10, min(28, max(data_max, header_len)))
            ws.set_column(i, i, width)

        # Ayarlar sayfası
        ayar_df = pd.DataFrame([
            ('Şablon', sablon_adi),
            ('Damızlık min puan', sablon['damizlik_min_puan']),
            ('Damızlık max yaş', sablon['damizlik_max_yas']),
            ('Reforme max puan', sablon['reforme_max_puan']),
            ('Reforme min yaş', sablon['reforme_min_yas']),
            ('Tek doğum cezası', sablon.get('tek_dogum_cezasi', False)),
            ('Tek doğum min yaş', sablon.get('tek_dogum_min_yas', '')),
            ('Tek doğum max puan', sablon.get('tek_dogum_max_puan', '')),
        ] + [(f'Ağırlık {y}', a) for y, a in zip(yillar, agirliklar)],
            columns=['Parametre', 'Değer'])
        ayar_df.to_excel(xw, sheet_name='Ayarlar', index=False)

    return bio.getvalue()


def sheets_yapistir_formati(df: pd.DataFrame, yillar: list[int]) -> str:
    """Google Sheets 'Puanlar' sayfasına yapıştırılacak tab-separated metin.
    Şema: Koyun No | Yıl Sayısı | Puan_YYYY... | Basit Ort | Ağırlıklı Ort | Rank Basit | Rank Ağırlıklı | Öneri
    """
    kols = ['Koyun No']
    if 'UKN' in df.columns:
        kols.append('UKN')
    kols += [f'Puan {y}' for y in yillar]
    kols += ['Basit Ort', 'Ağırlıklı Ort', 'Yıl Sayısı', 'Rank Basit', 'Rank Ağırlıklı', 'Yaş', 'Doğum Tipi', 'Öneri']
    kols = [k for k in kols if k in df.columns]
    return df[kols].to_csv(sep='\t', index=False, lineterminator='\n')


# ═══════════════════════════════════════════════════════════════════════════
# STREAMLIT UI
# ═══════════════════════════════════════════════════════════════════════════

st.set_page_config(page_title='Fenotipik Puan Birleştirme', layout='wide')

st.title('Koyun Fenotipik Puan Birleştirme')
st.caption(
    'Yıllık `Fenotipik_Indeks_*.xlsx` çıktılarını birleştir, '
    'çoklu yıl ortalama puan ve damızlık/reforme önerisi hesapla.'
)

# ──────────────────────────────────────────────────────────────────────────
# SIDEBAR — AYARLAR
# ──────────────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header('Ayarlar')

    st.subheader('1. Ağırlıklı Ortalama')
    st.caption('Son yıllar daha önemli — eksik yıl 0 sayılır.')
    agirlik_son  = st.number_input('Son yıl ağırlığı',  value=2.0, min_value=0.0, step=0.1)
    son_yil_sayisi = st.number_input('Kaç son yıl bu ağırlıkta?', value=1, min_value=1, step=1)
    agirlik_orta = st.number_input('Orta yıl ağırlığı', value=1.5, min_value=0.0, step=0.1)
    orta_yil_sayisi = st.number_input('Kaç orta yıl bu ağırlıkta?', value=1, min_value=1, step=1)
    agirlik_eski = st.number_input('Eski yıl ağırlığı', value=1.0, min_value=0.0, step=0.1)

    st.subheader('2. Öneri Şablonu')
    sablon_adi = st.selectbox(
        'Şablon',
        ['Standart', 'Sıkı', 'Gevşek', 'Özel'],
        index=0,
        help='Hibrit kural: Puan + Yaş + (opsiyonel) Tek doğum cezası',
    )

    if sablon_adi == 'Özel':
        if 'ozel_sablon' not in st.session_state:
            st.session_state.ozel_sablon = dict(SABLONLAR['Standart'])
        s = st.session_state.ozel_sablon
        c1, c2 = st.columns(2)
        s['damizlik_min_puan']  = c1.number_input('Damızlık min puan', value=float(s['damizlik_min_puan']), step=1.0)
        s['damizlik_max_yas']   = c2.number_input('Damızlık max yaş', value=int(s['damizlik_max_yas']), step=1)
        s['reforme_max_puan']   = c1.number_input('Reforme max puan', value=float(s['reforme_max_puan']), step=1.0)
        s['reforme_min_yas']    = c2.number_input('Reforme min yaş', value=int(s['reforme_min_yas']), step=1)
        s['tek_dogum_cezasi']   = st.checkbox('Tek doğum cezası', value=bool(s['tek_dogum_cezasi']))
        if s['tek_dogum_cezasi']:
            c1, c2 = st.columns(2)
            s['tek_dogum_min_yas']  = c1.number_input('TD min yaş',  value=int(s['tek_dogum_min_yas']), step=1)
            s['tek_dogum_max_puan'] = c2.number_input('TD max puan', value=float(s['tek_dogum_max_puan']), step=1.0)
        secili_sablon = s
    else:
        secili_sablon = SABLONLAR[sablon_adi]
        with st.expander(f'"{sablon_adi}" şablon detayları', expanded=False):
            st.json(secili_sablon)

    st.subheader('3. Öneri Hangi Ortalamaya Göre?')
    hangi_ort = st.radio(
        'Hangi ortalamaya göre öneri yapılsın?',
        ['Basit Ort', 'Ağırlıklı Ort'],
        index=0,
        horizontal=True,
    )

# ──────────────────────────────────────────────────────────────────────────
# ADIM 1 — YILLIK PUAN DOSYALARINI YÜKLE
# ──────────────────────────────────────────────────────────────────────────
st.header('1. Yıllık Fenotipik İndeks Dosyalarını Yükle')
st.caption('`Fenotipik_Indeks_App` çıktısı olan `Fenotipik_Indeks_YYYY*.xlsx` dosyaları.')

puan_files = st.file_uploader(
    'Birden fazla dosya seç',
    type=['xlsx'],
    accept_multiple_files=True,
    key='puan_files',
)

if not puan_files:
    st.info('En az bir yıllık puan dosyası yükle.')
    st.stop()

# Her dosya için yıl belirle (otomatik + override)
st.subheader('Dosya → Yıl Eşleştirme')
puan_dosyalari = {}
dosya_ozet = []
for i, f in enumerate(puan_files):
    auto_yil = yil_bul_dosya_adindan(f.name) or (datetime.now().year - len(puan_files) + i + 1)
    c1, c2, c3 = st.columns([3, 1, 2])
    c1.text(f.name)
    yil = c2.number_input(
        'Yıl', min_value=2000, max_value=2099,
        value=int(auto_yil), step=1, key=f'yil_{i}',
        label_visibility='collapsed',
    )
    try:
        df_y, _ = puan_dosyasi_oku(f, yil_override=int(yil))
        puan_dosyalari[int(yil)] = df_y
        c3.success(f'✓ {len(df_y)} koyun')
        dosya_ozet.append((int(yil), len(df_y), df_y['Puan'].mean()))
    except Exception as e:
        c3.error(f'Hata: {e}')

if not puan_dosyalari:
    st.error('Hiçbir dosya okunamadı.')
    st.stop()

# Aynı yıl iki kez yüklenmişse uyar
if len(puan_dosyalari) < len(puan_files):
    st.warning('Aynı yıl birden fazla dosya yüklenmiş — sonuncusu kullanıldı.')

# Yıl özet tablosu
if dosya_ozet:
    ozet_df = pd.DataFrame(dosya_ozet, columns=['Yıl', 'Koyun Sayısı', 'Ortalama Puan'])
    ozet_df['Ortalama Puan'] = ozet_df['Ortalama Puan'].round(2)
    st.dataframe(ozet_df.sort_values('Yıl'), use_container_width=True, hide_index=True)

# ──────────────────────────────────────────────────────────────────────────
# ADIM 2 — AKTİF KOYUN LİSTESİ (OPSİYONEL)
# ──────────────────────────────────────────────────────────────────────────
st.header('2. Aktif Koyun Listesi')
st.caption(
    'Çiftleşmesi muhtemel, şu an yaşayan koyunların listesi (Excel/CSV). '
    'Yüklenmezse tüm puan dosyalarındaki koyunlar gösterilir. '
    "Sütun isimleri: 'Koyun No' zorunlu; opsiyonel: UKN, Yaş, Doğum Tipi, Padok, Irk."
)

aktif_file = st.file_uploader(
    'Aktif liste dosyası (Excel/CSV) — opsiyonel',
    type=['xlsx', 'csv'],
    key='aktif_file',
)

aktif_df = None
if aktif_file:
    try:
        aktif_df = aktif_liste_oku(aktif_file)
        st.success(f'✓ Aktif listede {len(aktif_df)} koyun.')
        with st.expander('Aktif liste önizleme', expanded=False):
            st.dataframe(aktif_df.head(20), use_container_width=True, hide_index=True)
    except Exception as e:
        st.error(f'Aktif liste okunamadı: {e}')

# ──────────────────────────────────────────────────────────────────────────
# ADIM 3 — BİRLEŞTİR + HESAPLA
# ──────────────────────────────────────────────────────────────────────────
st.header('3. Birleştir ve Hesapla')

birlesim, yillar = birlestir(puan_dosyalari, aktif_df)
if len(birlesim) == 0:
    st.error('Birleştirme sonucu boş.')
    st.stop()

birlesim, agirliklar = ortalamalari_hesapla(
    birlesim, yillar,
    agirlik_son=agirlik_son, agirlik_orta=agirlik_orta, agirlik_eski=agirlik_eski,
    son_yil_sayisi=int(son_yil_sayisi), orta_yil_sayisi=int(orta_yil_sayisi),
)
birlesim['Öneri'] = oneri_uret(birlesim, secili_sablon, hangi_ort=hangi_ort)

# Sıralama
birlesim_sorted = birlesim.sort_values(hangi_ort, ascending=False).reset_index(drop=True)

# Özet metrikler
c1, c2, c3, c4 = st.columns(4)
c1.metric('Toplam Koyun', len(birlesim_sorted))
c2.metric('Damızlık Önerisi', int((birlesim_sorted['Öneri'] == 'Damızlık').sum()))
c3.metric('Reforme Önerisi',  int((birlesim_sorted['Öneri'] == 'Reforme').sum()))
c4.metric('Sınırda',           int((birlesim_sorted['Öneri'] == 'Sınırda').sum()))

# Ağırlık özeti
ag_yazi = ', '.join(f'{y}→{a:.1f}' for y, a in zip(yillar, agirliklar))
st.caption(f'Ağırlıklar: {ag_yazi}')

# Sütun düzeni
goster_kols = ['Koyun No']
for opt in ('UKN', 'Padok', 'Irk'):
    if opt in birlesim_sorted.columns:
        goster_kols.append(opt)
goster_kols += ['Yaş', 'Doğum Tipi']
goster_kols += [f'Puan {y}' for y in yillar]
goster_kols += ['Basit Ort', 'Ağırlıklı Ort', 'Yıl Sayısı', 'Rank Basit', 'Rank Ağırlıklı', 'Öneri']
goster_kols = [k for k in goster_kols if k in birlesim_sorted.columns]

# Renklendirme
def renklendir_oneri(val):
    if val == 'Damızlık':
        return 'background-color: #d4edda; color: #155724; font-weight: bold'
    if val == 'Reforme':
        return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
    return 'background-color: #fff3cd; color: #856404'

st.subheader('Sonuç Tablosu (makrolu Excel "Puan" sayfası karşılığı)')
styled = birlesim_sorted[goster_kols].style.map(renklendir_oneri, subset=['Öneri'])
st.dataframe(styled, use_container_width=True, height=600)

# Filtre kutusu — hızlı arama
with st.expander('🔍 Filtre / Arama', expanded=False):
    c1, c2, c3 = st.columns(3)
    oneri_filtre = c1.multiselect('Öneri', ['Damızlık', 'Reforme', 'Sınırda'], default=[])
    min_puan = c2.number_input('Min ortalama puan', value=0.0, step=1.0)
    arama_kn = c3.text_input('Koyun No ara')
    f = birlesim_sorted.copy()
    if oneri_filtre:
        f = f[f['Öneri'].isin(oneri_filtre)]
    f = f[f[hangi_ort] >= min_puan]
    if arama_kn.strip():
        f = f[f['Koyun No'].astype(str).str.contains(arama_kn.strip())]
    st.dataframe(
        f[goster_kols].style.map(renklendir_oneri, subset=['Öneri']),
        use_container_width=True, height=400,
    )
    st.caption(f'{len(f)} sonuç.')

# ──────────────────────────────────────────────────────────────────────────
# ADIM 4 — İNDİR
# ──────────────────────────────────────────────────────────────────────────
st.header('4. İndir / Sheets\'e Yapıştır')

ts = datetime.now().strftime('%Y%m%d_%H%M')
c1, c2, c3 = st.columns(3)

with c1:
    xlsx_bytes = excel_olarak_indir(
        birlesim_sorted[goster_kols], yillar, agirliklar, secili_sablon, sablon_adi,
    )
    st.download_button(
        'Excel olarak indir (.xlsx)',
        data=xlsx_bytes,
        file_name=f'Puan_Birlesim_{ts}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        type='primary',
    )

with c2:
    csv_bytes = birlesim_sorted[goster_kols].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button(
        'CSV olarak indir (.csv)',
        data=csv_bytes,
        file_name=f'Puan_Birlesim_{ts}.csv',
        mime='text/csv',
    )

with c3:
    sheets_metni = sheets_yapistir_formati(birlesim_sorted, yillar)
    st.download_button(
        'Sheets için tab-delimited (.tsv)',
        data=sheets_metni.encode('utf-8'),
        file_name=f'Puanlar_Sheets_{ts}.tsv',
        mime='text/tab-separated-values',
        help='Google Sheets\'te "Puanlar" sayfasına yapıştırmak için.',
    )

with st.expander('Sheets\'e yapıştırma talimatı', expanded=False):
    st.markdown("""
1. Yukarıdaki **'Sheets için tab-delimited'** butonuna bas → `.tsv` indir.
2. Google Sheets'te "Puanlar" sayfasını aç.
3. A1 hücresine tıkla, dosyadaki tüm metni kopyala (Ctrl+A → Ctrl+C), Sheets'e yapıştır (Ctrl+V).
4. Veya doğrudan aşağıdaki kutudan kopyala-yapıştır yapabilirsin:
""")
    st.code(sheets_metni[:5000] + ('\n... (kısaltıldı, dosyayı indir)' if len(sheets_metni) > 5000 else ''),
            language='text')
