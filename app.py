import os
import re
import html
import pandas as pd
import streamlit as st

# =====================================================
# KONFIGURASI
# =====================================================
DEFAULT_DATASET = "Data set 20098+ Gambar.xlsx"

st.set_page_config(
    page_title="HyTBIONEX",
    page_icon="🌿",
    layout="wide"
)


# =====================================================
# CSS TAMPILAN
# =====================================================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #e8f8ec 0%, #f3e8ff 100%);
}

[data-testid="stSidebar"] {
    background: #262832;
}

[data-testid="stSidebar"] * {
    color: white !important;
}

.main-title {
    background: linear-gradient(135deg, #0b7a45, #129157);
    padding: 35px;
    border-radius: 25px;
    text-align: center;
    color: white;
    margin-bottom: 25px;
    box-shadow: 0 8px 20px rgba(0,0,0,0.18);
}

.main-title h1 {
    font-size: 54px;
    font-weight: 900;
    margin-bottom: 8px;
}

.orange-card {
    background: linear-gradient(135deg, #f97316, #fb923c);
    padding: 25px;
    border-radius: 22px;
    color: white;
    margin-bottom: 25px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.15);
}

.green-card {
    background: linear-gradient(135deg, #0b7a45, #129157);
    padding: 25px;
    border-radius: 22px;
    color: white;
    margin-bottom: 15px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.15);
}

.lilac-card {
    background: #f3e8ff;
    padding: 22px;
    border-radius: 20px;
    border: 2px solid #c084fc;
    color: #111111;
    margin-top: 18px;
    margin-bottom: 18px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.12);
}

.result-box {
    background: white;
    border-left: 8px solid #0b7a45;
    padding: 18px;
    border-radius: 16px;
    margin-bottom: 12px;
    color: #111111;
    box-shadow: 0 6px 14px rgba(0,0,0,0.10);
}

.relation-box {
    background: #fff7ed;
    border: 2px solid #fb923c;
    padding: 18px;
    border-radius: 16px;
    margin-bottom: 12px;
    color: #111111;
}

.kg-box {
    background: radial-gradient(circle, #f3e8ff, #ffffff);
    border: 2px dashed #a855f7;
    padding: 22px;
    border-radius: 20px;
    margin-top: 18px;
    color: #111111;
}

.node {
    display: inline-block;
    padding: 12px 18px;
    border-radius: 999px;
    margin: 8px;
    font-weight: 800;
    color: #111111;
}

.node-main {
    background: #86efac;
    border: 3px solid #0b7a45;
}

.node-lilac {
    background: #e9d5ff;
    border: 2px solid #a855f7;
}

.node-orange {
    background: #fed7aa;
    border: 2px solid #f97316;
}

.node-green {
    background: #bbf7d0;
    border: 2px solid #16a34a;
}

.small-note {
    font-size: 14px;
    color: #333333;
}
</style>
""", unsafe_allow_html=True)


# =====================================================
# FUNGSI BANTUAN
# =====================================================
def clean_text(text):
    text = str(text).lower()
    text = re.sub(r"[^a-zA-Z0-9À-ÿ\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def safe(text):
    if text is None:
        return ""
    return html.escape(str(text))


def normalize_colname(text):
    text = str(text).strip().lower()
    text = text.replace("_", " ")
    text = text.replace("/", " ")
    text = re.sub(r"[^a-zA-Z0-9À-ÿ\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def get_col(row, possible_names):
    for wanted in possible_names:
        wanted_norm = normalize_colname(wanted)
        for col in row.index:
            if normalize_colname(col) == wanted_norm:
                value = str(row[col]).strip()
                if value.lower() not in ["", "nan", "none"]:
                    return value
    return ""


def find_dataset_file():
    if os.path.exists(DEFAULT_DATASET):
        return DEFAULT_DATASET

    excel_files = []
    for f in os.listdir("."):
        if f.lower().endswith((".xlsx", ".xls")):
            excel_files.append(f)

    for f in excel_files:
        if "data" in f.lower():
            return f

    if excel_files:
        return excel_files[0]

    return ""


@st.cache_data
def load_dataset():
    dataset_path = find_dataset_file()

    if not dataset_path:
        return pd.DataFrame(), "Dataset Excel belum ditemukan di GitHub."

    try:
        sheets = pd.read_excel(dataset_path, sheet_name=None)
        frames = []

        for sheet_name, df in sheets.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                df = df.fillna("")
                df["__sheet_name__"] = sheet_name
                frames.append(df)

        if frames:
            data = pd.concat(frames, ignore_index=True).fillna("")
            return data, f"Dataset terbaca: {dataset_path} | Total data: {len(data)} baris"

        return pd.DataFrame(), f"Dataset kosong: {dataset_path}"

    except Exception as e:
        return pd.DataFrame(), f"Gagal membaca dataset: {e}"


def read_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return "", "Tidak ada dokumen yang diupload."

    name = uploaded_file.name.lower()

    try:
        if name.endswith(".pdf"):
            from pypdf import PdfReader
            reader = PdfReader(uploaded_file)
            text = ""
            for page in reader.pages:
                text += " " + (page.extract_text() or "")
            if not text.strip():
                return "", f"PDF terbaca tetapi teks kosong: {uploaded_file.name}. Kemungkinan PDF berbentuk scan/gambar."
            return text, f"PDF terbaca: {uploaded_file.name} ({len(text)} karakter)"

        elif name.endswith(".txt"):
            text = uploaded_file.read().decode("utf-8", errors="ignore")
            return text, f"TXT terbaca: {uploaded_file.name} ({len(text)} karakter)"

        elif name.endswith(".csv"):
            df = pd.read_csv(uploaded_file).fillna("")
            text = " ".join(df.astype(str).values.flatten())
            return text, f"CSV terbaca: {uploaded_file.name} ({len(text)} karakter)"

        elif name.endswith((".xlsx", ".xls")):
            sheets = pd.read_excel(uploaded_file, sheet_name=None)
            text = ""
            for _, df in sheets.items():
                df = df.fillna("")
                text += " " + " ".join(df.astype(str).values.flatten())
            return text, f"Excel dokumen terbaca: {uploaded_file.name} ({len(text)} karakter)"

        else:
            return "", "Format dokumen belum didukung."

    except Exception as e:
        return "", f"Gagal membaca dokumen: {e}"


def score_match(row, search_text):
    nama_tanaman = clean_text(get_col(row, [
        "Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"
    ]))

    nama_latin = clean_text(get_col(row, [
        "Nama Latin", "Nama_Latin", "Latin"
    ]))

    nama_lokal = clean_text(get_col(row, [
        "Nama Lokal/Daerah", "Nama Lokal", "Nama_Daerah", "Bahasa_Daerah", "Bahasa Daerah"
    ]))

    score = 0

    if nama_tanaman and nama_tanaman in search_text:
        score += 150

    if nama_latin and nama_latin in search_text:
        score += 150

    if nama_lokal:
        for p in re.split(r"[,;/|]", nama_lokal):
            p = clean_text(p)
            if p and len(p) >= 3 and p in search_text:
                score += 80

    return score


def find_best_row(dataset_df, search_text):
    if dataset_df.empty:
        return None, "Dataset belum terbaca."

    if not search_text:
        return None, "Input tanaman dan dokumen masih kosong."

    best_row = None
    best_score = 0

    for _, row in dataset_df.iterrows():
        score = score_match(row, search_text)
        if score > best_score:
            best_score = score
            best_row = row

    if best_row is not None and best_score > 0:
        nama = get_col(best_row, ["Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"])
        latin = get_col(best_row, ["Nama Latin", "Nama_Latin", "Latin"])
        return best_row, f"Entitas cocok dengan dataset: {nama} / {latin} | Skor: {best_score}"

    return None, "Tidak ditemukan kecocokan entitas tanaman pada dataset."


def make_result(row, input_text):
    if row is None:
        return {
            "Nama Tanaman": input_text if input_text else "Belum terdeteksi",
            "Nama Latin": "Belum terdeteksi",
            "Nama Lokal/Daerah": "Belum terdeteksi",
            "Bagian Tanaman": "Belum terdeteksi",
            "Zat Bioaktif": "Belum terdeteksi",
            "Khasiat/Efek Terapeutik": "Belum terdeteksi",
            "Cara Pengolahan": "Belum terdeteksi",
            "Komposisi/Dosis": "Belum terdeteksi",
            "Sumber Data": "Belum terdeteksi"
        }

    return {
        "Nama Tanaman": get_col(row, [
            "Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"
        ]),
        "Nama Latin": get_col(row, [
            "Nama Latin", "Nama_Latin", "Latin"
        ]),
        "Nama Lokal/Daerah": get_col(row, [
            "Nama Lokal/Daerah", "Nama Lokal", "Nama_Daerah",
            "Bahasa_Daerah", "Bahasa Daerah"
        ]),
        "Bagian Tanaman": get_col(row, [
            "Bagian Tanaman", "Bagian_Tanaman", "Bagian Digunakan",
            "Bagian_Digunakan", "Bagian"
        ]),
        "Zat Bioaktif": get_col(row, [
            "Zat Bioaktif", "Senyawa Bioaktif", "Senyawa_Bioaktif",
            "Compound", "Senyawa", "Kandungan", "Kandungan Kimia"
        ]),
        "Khasiat/Efek Terapeutik": get_col(row, [
            "Khasiat/Efek Terapeutik", "Khasiat", "Benefit",
            "Biological_Activity", "Aktivitas Farmakologis", "Manfaat"
        ]),
        "Cara Pengolahan": get_col(row, [
            "Cara Pengolahan", "Cara_Pengolahan", "Pengolahan", "Cara Pemakaian"
        ]),
        "Komposisi/Dosis": get_col(row, [
            "Komposisi/Dosis", "Komposisi /Dosis", "Dosis", "Komposisi"
        ]),
        "Sumber Data": get_col(row, [
            "Sumber Data", "Sumber_Data", "Sumber", "Referensi"
        ])
    }


def find_image(row):
    if row is None:
        return None

    img_name = get_col(row, [
        "Gambar", "gambar", "Image", "image", "Foto", "foto",
        "Nama File Gambar", "File Gambar", "Path Gambar"
    ])

    if not img_name:
        return None

    candidates = [
        img_name,
        f"assets/{img_name}",
        f"gambar/{img_name}",
        f"images/{img_name}",
        f"foto/{img_name}",
        f"assets/{img_name}.png",
        f"assets/{img_name}.jpg",
        f"assets/{img_name}.jpeg",
        f"gambar/{img_name}.png",
        f"gambar/{img_name}.jpg",
        f"gambar/{img_name}.jpeg",
    ]

    for path in candidates:
        if os.path.exists(path):
            return path

    return None


def show_result_card(icon, title, value):
    st.markdown(
        f"""
        <div class="result-box">
            <h4>{icon} {safe(title)}</h4>
            <p>{safe(value) if value else "Belum terdeteksi"}</p>
        </div>
        """,
        unsafe_allow_html=True
    )


# =====================================================
# SIDEBAR
# =====================================================
st.sidebar.markdown("# 🌱 HyTBIONEX")
st.sidebar.markdown("---")
st.sidebar.markdown("🏠 Dashboard Utama")
st.sidebar.markdown("🌿 Input Data Tanaman")
st.sidebar.markdown("📁 Upload Dokumen")
st.sidebar.markdown("📋 Hasil Ekstraksi")
st.sidebar.markdown("🔗 Relation Extraction")
st.sidebar.markdown("🕸️ HerbKG 2.0")
st.sidebar.markdown("---")
st.sidebar.markdown("### Advanced Downstream Applications")
st.sidebar.markdown("📊 Descriptive Analytics")
st.sidebar.markdown("🔎 Evidence-Based Graph Query")
st.sidebar.markdown("🧬 Similarity Analysis")
st.sidebar.markdown("💊 Herbal Recommendation")


# =====================================================
# HEADER
# =====================================================
st.markdown("""
<div class="main-title">
    <h1>🌿 HyTBIONEX</h1>
    <h3>Hybrid Transformer for Bioactive Information Extraction</h3>
    <p>Analisis Bioaktif Tanaman Herbal Indonesia & Enhanced Herb Knowledge Graph 2.0</p>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="orange-card">
    <h2>Selamat Datang di HyTBIONEX</h2>
    <p>
    Platform cerdas untuk ekstraksi informasi bioaktif tanaman herbal Indonesia
    berbasis Hybrid Transformer, Bioactive Information Extraction,
    Named Entity Disambiguation, Relation Extraction, dan Enhanced Herb Knowledge Graph.
    </p>
</div>
""", unsafe_allow_html=True)


# =====================================================
# STATUS DATASET
# =====================================================
dataset_df, dataset_status = load_dataset()

st.markdown(
    f"""
    <div class="lilac-card">
        <h3>📌 Status Dataset</h3>
        <p>{safe(dataset_status)}</p>
    </div>
    """,
    unsafe_allow_html=True
)


# =====================================================
# INPUT AREA
# =====================================================
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="green-card">
        <h2>🌿 1. Input Data Tanaman</h2>
        <p>Masukkan nama tanaman atau kalimat artikel herbal.</p>
    </div>
    """, unsafe_allow_html=True)

    teks = st.text_area(
        "Input Data Tanaman",
        placeholder="Contoh: Kelor, Sirih, Jahe, Kunyit, atau kalimat artikel herbal",
        height=180
    )

with col2:
    st.markdown("""
    <div class="green-card">
        <h2>📁 2. Upload Dokumen Artikel / Dataset</h2>
        <p>Upload PDF, TXT, CSV, atau Excel.</p>
    </div>
    """, unsafe_allow_html=True)

    dokumen = st.file_uploader(
        "Upload Dokumen",
        type=["pdf", "txt", "csv", "xlsx", "xls"]
    )


# =====================================================
# PROSES EKSTRAKSI
# =====================================================
if st.button("🔍 PROSES EKSTRAKSI", use_container_width=True):
    progress = st.progress(0)
    status = st.empty()

    status.text("Memulai proses ekstraksi... 10%")
    progress.progress(10)

    status.text("Membaca input tanaman... 30%")
    progress.progress(30)

    doc_text, doc_status = read_uploaded_file(dokumen)
    status.text("Membaca dokumen... 50%")
    progress.progress(50)

    search_text = clean_text(teks) + " " + clean_text(doc_text)
    status.text("Mencocokkan entitas dengan dataset... 75%")
    progress.progress(75)

    best_row, match_status = find_best_row(dataset_df, search_text)
    result = make_result(best_row, teks)
    image_path = find_image(best_row)

    status.text("Membangun hasil ekstraksi... 100%")
    progress.progress(100)

    st.success("Proses ekstraksi selesai.")

    st.markdown(
        f"""
        <div class="lilac-card">
            <h3>📌 Status Sistem</h3>
            <p><b>Status Dataset:</b> {safe(dataset_status)}</p>
            <p><b>Status Dokumen:</b> {safe(doc_status)}</p>
            <p><b>Status Koneksi Entitas:</b> {safe(match_status)}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("## 📋 Hasil Ekstraksi Informasi Bioaktif")

    c1, c2, c3 = st.columns(3)
    with c1:
        show_result_card("🌿", "Nama Tanaman", result["Nama Tanaman"])
        show_result_card("🍃", "Bagian Tanaman", result["Bagian Tanaman"])
        show_result_card("☕", "Cara Pengolahan", result["Cara Pengolahan"])

    with c2:
        show_result_card("🔬", "Nama Latin", result["Nama Latin"])
        show_result_card("🧪", "Zat Bioaktif", result["Zat Bioaktif"])
        show_result_card("⚖️", "Komposisi/Dosis", result["Komposisi/Dosis"])

    with c3:
        show_result_card("🇮🇩", "Nama Lokal/Daerah", result["Nama Lokal/Daerah"])
        show_result_card("💚", "Khasiat/Efek Terapeutik", result["Khasiat/Efek Terapeutik"])
        show_result_card("📚", "Sumber Data", result["Sumber Data"])

    st.markdown("## 🖼️ Lampiran Gambar Tanaman")
    if image_path:
        st.image(image_path, caption="Gambar tanaman dari metadata dataset", use_container_width=True)
    else:
        st.info("Gambar tanaman belum tersedia. Kolom gambar/path gambar belum ditemukan atau file gambar belum diupload.")

    st.markdown("## 🔗 Bioactive Relation Extraction")
    st.markdown(
        f"""
        <div class="relation-box">
            <p><b>{safe(result["Nama Tanaman"])}</b> → has_latin_name → <b>{safe(result["Nama Latin"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → has_local_name → <b>{safe(result["Nama Lokal/Daerah"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → uses_part → <b>{safe(result["Bagian Tanaman"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → contains_bioactive_compound → <b>{safe(result["Zat Bioaktif"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → has_therapeutic_effect → <b>{safe(result["Khasiat/Efek Terapeutik"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → processed_by → <b>{safe(result["Cara Pengolahan"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → has_dosage → <b>{safe(result["Komposisi/Dosis"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → sourced_from → <b>{safe(result["Sumber Data"])}</b></p>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown("## 🕸️ Enhanced Herb Knowledge Graph 2.0")
    st.markdown(
        f"""
        <div class="kg-box">
            <div style="text-align:center;">
                <span class="node node-main">🌿 {safe(result["Nama Tanaman"])}</span>
            </div>
            <div style="text-align:center; margin-top:15px;">
                <span class="node node-lilac">🔬 {safe(result["Nama Latin"])}</span>
                <span class="node node-green">🍃 {safe(result["Bagian Tanaman"])}</span>
                <span class="node node-orange">🧪 {safe(result["Zat Bioaktif"])}</span>
                <span class="node node-lilac">💚 {safe(result["Khasiat/Efek Terapeutik"])}</span>
                <span class="node node-green">☕ {safe(result["Cara Pengolahan"])}</span>
                <span class="node node-orange">⚖️ {safe(result["Komposisi/Dosis"])}</span>
                <span class="node node-lilac">📚 {safe(result["Sumber Data"])}</span>
            </div>
            <p class="small-note">
            Visualisasi ini menunjukkan representasi awal HerbKG 2.0 berbasis hasil ekstraksi entitas dan relasi.
            </p>
        </div>
        """,
        unsafe_allow_html=True
    )

else:
    st.info("Masukkan nama tanaman atau upload dokumen, lalu klik tombol PROSES EKSTRAKSI.")
