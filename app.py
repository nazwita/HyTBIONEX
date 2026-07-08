import os
import re
import html
import base64
import textwrap
import pandas as pd
import streamlit as st
import networkx as nx
import matplotlib.pyplot as plt

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
# SESSION STATE
# =====================================================
if "last_result" not in st.session_state:
    st.session_state.last_result = None

if "last_dataset_status" not in st.session_state:
    st.session_state.last_dataset_status = ""

if "last_doc_status" not in st.session_state:
    st.session_state.last_doc_status = ""

if "last_match_status" not in st.session_state:
    st.session_state.last_match_status = ""

if "input_text_area" not in st.session_state:
    st.session_state.input_text_area = ""

if "file_key_suffix" not in st.session_state:
    st.session_state.file_key_suffix = 0


def clear_input_text():
    st.session_state.input_text_area = ""


def clear_uploaded_file():
    st.session_state.file_key_suffix += 1
    st.session_state.input_text_area = ""


def clear_text_when_upload():
    st.session_state.input_text_area = ""


# =====================================================
# BACKGROUND IMAGE OPSIONAL
# =====================================================
def get_image_data_url(paths):
    for path in paths:
        if os.path.exists(path):
            try:
                with open(path, "rb") as f:
                    encoded = base64.b64encode(f.read()).decode()
                ext = path.split(".")[-1].lower()
                if ext == "jpg":
                    ext = "jpeg"
                return f"data:image/{ext};base64,{encoded}"
            except Exception:
                pass
    return ""


BG_URL = get_image_data_url([
    "assets/background.png",
    "assets/herbal-background.png",
    "background.png",
    "ChatGPT Image 9 Jul 2026, 05.26.27.png"
])

if BG_URL:
    HERO_BG = f"linear-gradient(90deg, rgba(255,255,255,0.92), rgba(255,255,255,0.40)), url('{BG_URL}')"
else:
    HERO_BG = "radial-gradient(circle at 70% 30%, #d9f99d 0%, #ecfdf5 35%, #fffaf0 100%)"


# =====================================================
# CSS
# =====================================================
st.markdown(f"""
<style>
.stApp {{
    background:
        radial-gradient(circle at top left, rgba(187,247,208,0.55), transparent 32%),
        radial-gradient(circle at bottom right, rgba(233,213,255,0.75), transparent 35%),
        linear-gradient(135deg, #f8fff8 0%, #fffaf0 45%, #f3e8ff 100%) !important;
}}

/* SIDEBAR */
[data-testid="stSidebar"] {{
    background:
        radial-gradient(circle at bottom left, rgba(255,237,213,0.14), transparent 30%),
        linear-gradient(180deg, #013220 0%, #064e3b 45%, #047857 100%) !important;
    border-right: 1px solid rgba(255,255,255,0.12);
}}

[data-testid="stSidebar"] * {{
    color: #ffedd5 !important;
}}

[data-testid="stSidebar"] h1 {{
    color: #ffffff !important;
    font-size: 34px !important;
    font-weight: 900 !important;
}}

[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {{
    color: #fed7aa !important;
    font-weight: 900 !important;
}}

[data-testid="stSidebar"] [role="radiogroup"] label {{
    background: rgba(255, 237, 213, 0.10) !important;
    padding: 13px 15px !important;
    border-radius: 15px !important;
    margin-bottom: 9px !important;
    border: 1px solid rgba(255, 237, 213, 0.14) !important;
}}

[data-testid="stSidebar"] [role="radiogroup"] label:hover {{
    background: rgba(249, 115, 22, 0.32) !important;
    transform: translateX(4px);
}}

/* HEADER MOCKUP */
.top-header {{
    background:
        linear-gradient(135deg, rgba(1,50,32,0.98), rgba(6,78,59,0.96), rgba(4,120,87,0.92));
    padding: 26px 32px;
    border-radius: 26px;
    color: white;
    margin-bottom: 22px;
    box-shadow: 0 18px 40px rgba(0,0,0,0.18);
    display: flex;
    align-items: center;
    justify-content: space-between;
}}

.top-header h1 {{
    margin: 0;
    font-size: 42px;
    font-weight: 900;
}}

.top-header p {{
    margin: 4px 0 0 0;
    font-size: 17px;
    color: #ecfdf5;
}}

.status-pill {{
    background: rgba(255,255,255,0.10);
    padding: 12px 18px;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.14);
    font-weight: 800;
    color: #ffedd5;
}}

/* HERO BOTANICAL */
.hero {{
    background: {HERO_BG};
    background-size: cover;
    background-position: center;
    padding: 34px;
    border-radius: 28px;
    margin-bottom: 22px;
    box-shadow: 0 18px 38px rgba(0,0,0,0.12);
    border: 1px solid rgba(6,78,59,0.10);
}}

.hero h1 {{
    color: #064e3b;
    font-size: 56px;
    font-weight: 900;
    margin-bottom: 6px;
}}

.hero h2 {{
    color: #0b7a45;
    font-size: 25px;
    font-weight: 900;
}}

.hero p {{
    color: #12372a;
    font-size: 17px;
    line-height: 1.7;
}}

.preview-card {{
    background: rgba(255,255,255,0.82);
    border: 1px solid rgba(6,78,59,0.14);
    border-radius: 24px;
    padding: 22px;
    box-shadow: 0 12px 26px rgba(0,0,0,0.10);
}}

.orange-card {{
    background: linear-gradient(135deg, #f97316, #fb923c);
    padding: 25px;
    border-radius: 22px;
    color: white;
    margin-bottom: 25px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.15);
}}

.green-card {{
    background: linear-gradient(135deg, #064e3b, #0b7a45, #059669);
    padding: 24px;
    border-radius: 22px;
    color: white;
    margin-bottom: 15px;
    box-shadow: 0 14px 28px rgba(0,0,0,0.16);
    position: relative;
}}

.green-card h2 {{
    margin: 0 0 8px 0;
    color: white;
}}

.green-card p {{
    margin: 0;
    color: #ecfdf5;
}}

.lilac-card {{
    background: #f3e8ff;
    padding: 22px;
    border-radius: 20px;
    border: 2px solid #c084fc;
    color: #111111;
    margin-top: 18px;
    margin-bottom: 18px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.10);
}}

.white-card {{
    background: rgba(255,255,255,0.90);
    padding: 22px;
    border-radius: 22px;
    border: 1px solid rgba(6,78,59,0.10);
    margin-bottom: 18px;
    box-shadow: 0 10px 24px rgba(0,0,0,0.10);
}}

.result-box {{
    background: white;
    border-left: 8px solid #0b7a45;
    padding: 18px;
    border-radius: 16px;
    margin-bottom: 12px;
    color: #111111;
    box-shadow: 0 6px 14px rgba(0,0,0,0.10);
}}

.relation-box {{
    background: #fff7ed;
    border: 2px solid #fb923c;
    padding: 18px;
    border-radius: 16px;
    margin-bottom: 12px;
    color: #111111;
}}

.kg-box {{
    background: radial-gradient(circle, #f3e8ff, #ffffff);
    border: 3px dashed #a855f7;
    padding: 22px;
    border-radius: 20px;
    margin-top: 18px;
    color: #111111;
}}

/* INPUT DAN UPLOAD JADI LILAC, BUKAN HITAM */
textarea {{
    background-color: #f3e8ff !important;
    color: #111111 !important;
    border: 2px solid #c084fc !important;
    border-radius: 16px !important;
}}

textarea::placeholder {{
    color: #6b4b84 !important;
}}

[data-testid="stFileUploader"] section {{
    background-color: #f3e8ff !important;
    border: 2px dashed #a855f7 !important;
    border-radius: 16px !important;
    color: #111111 !important;
}}

[data-testid="stFileUploader"] section * {{
    color: #111111 !important;
}}

[data-testid="stFileUploader"] button {{
    background: #a855f7 !important;
    color: white !important;
    border-radius: 12px !important;
    border: none !important;
    font-weight: 800 !important;
}}

.stButton > button {{
    background: linear-gradient(135deg, #f97316, #fb923c) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    font-weight: 900 !important;
    padding: 0.75rem 1rem !important;
}}

.small-x button {{
    width: 34px !important;
    height: 34px !important;
    min-width: 34px !important;
    border-radius: 50% !important;
    padding: 0 !important;
    font-size: 18px !important;
    background: #fb923c !important;
    color: white !important;
    border: 2px solid white !important;
    margin-top: 3px !important;
}}

.metric-card {{
    background: rgba(255,255,255,0.90);
    padding: 17px;
    border-radius: 18px;
    border: 1px solid rgba(6,78,59,0.10);
    box-shadow: 0 8px 18px rgba(0,0,0,0.08);
}}

.metric-card h4 {{
    margin: 0;
    color: #047857;
}}

.metric-card h2 {{
    margin: 5px 0 0 0;
    color: #16a34a;
}}

.down-card {{
    background: rgba(255,255,255,0.92);
    border-radius: 22px;
    padding: 20px;
    border: 1px solid rgba(6,78,59,0.12);
    box-shadow: 0 12px 26px rgba(0,0,0,0.10);
    min-height: 310px;
}}

.down-card h3 {{
    color: #064e3b;
    font-size: 19px;
    font-weight: 900;
}}

.down-card p {{
    color: #12372a;
    font-size: 14px;
}}

.mini-table {{
    width: 100%;
    border-collapse: collapse;
    font-size: 12px;
}}

.mini-table td, .mini-table th {{
    border: 1px solid #d1d5db;
    padding: 5px;
    color: #111;
}}

.mini-flow {{
    background: #fff7ed;
    border: 1px solid #fdba74;
    padding: 8px;
    border-radius: 10px;
    text-align: center;
    margin-bottom: 7px;
    color: #111;
    font-size: 13px;
    font-weight: 700;
}}

.small-note {{
    font-size: 15px;
    color: #333333;
}}
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

    excel_files = [f for f in os.listdir(".") if f.lower().endswith((".xlsx", ".xls"))]

    for f in excel_files:
        if "data" in f.lower():
            return f

    return excel_files[0] if excel_files else ""


@st.cache_data(show_spinner=False)
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
                return "", f"PDF terbaca tetapi teks kosong: {uploaded_file.name}. Kemungkinan PDF scan/gambar."

            return text, f"PDF terbaca: {uploaded_file.name} ({len(text)} karakter)"

        if name.endswith(".txt"):
            text = uploaded_file.read().decode("utf-8", errors="ignore")
            return text, f"TXT terbaca: {uploaded_file.name} ({len(text)} karakter)"

        if name.endswith(".csv"):
            df = pd.read_csv(uploaded_file).fillna("")
            text = " ".join(df.astype(str).values.flatten())
            return text, f"CSV terbaca: {uploaded_file.name} ({len(text)} karakter)"

        if name.endswith((".xlsx", ".xls")):
            sheets = pd.read_excel(uploaded_file, sheet_name=None)
            text = ""
            for _, df in sheets.items():
                df = df.fillna("")
                text += " " + " ".join(df.astype(str).values.flatten())
            return text, f"Excel dokumen terbaca: {uploaded_file.name} ({len(text)} karakter)"

        return "", "Format dokumen belum didukung."

    except Exception as e:
        return "", f"Gagal membaca dokumen: {e}"


def score_match(row, search_text):
    nama_tanaman = clean_text(get_col(row, ["Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"]))
    nama_latin = clean_text(get_col(row, ["Nama Latin", "Nama_Latin", "Latin"]))
    nama_lokal = clean_text(get_col(row, ["Nama Lokal/Daerah", "Nama Lokal", "Nama_Daerah", "Bahasa_Daerah", "Bahasa Daerah"]))

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
        "Nama Tanaman": get_col(row, ["Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"]),
        "Nama Latin": get_col(row, ["Nama Latin", "Nama_Latin", "Latin"]),
        "Nama Lokal/Daerah": get_col(row, ["Nama Lokal/Daerah", "Nama Lokal", "Nama_Daerah", "Bahasa_Daerah", "Bahasa Daerah"]),
        "Bagian Tanaman": get_col(row, ["Bagian Tanaman", "Bagian_Tanaman", "Bagian Digunakan", "Bagian_Digunakan", "Bagian"]),
        "Zat Bioaktif": get_col(row, ["Zat Bioaktif", "Senyawa Bioaktif", "Senyawa_Bioaktif", "Compound", "Senyawa", "Kandungan", "Kandungan Kimia"]),
        "Khasiat/Efek Terapeutik": get_col(row, ["Khasiat/Efek Terapeutik", "Khasiat", "Benefit", "Biological_Activity", "Aktivitas Farmakologis", "Manfaat"]),
        "Cara Pengolahan": get_col(row, ["Cara Pengolahan", "Cara_Pengolahan", "Pengolahan", "Cara Pemakaian"]),
        "Komposisi/Dosis": get_col(row, ["Komposisi/Dosis", "Komposisi /Dosis", "Dosis", "Komposisi"]),
        "Sumber Data": get_col(row, ["Sumber Data", "Sumber_Data", "Sumber", "Referensi"])
    }


def run_extraction(teks, dokumen, dataset_df, dataset_status):
    progress = st.progress(0)
    status = st.empty()

    status.text("Memulai proses ekstraksi... 10%")
    progress.progress(10)

    status.text("Membaca input dan dokumen... 40%")
    doc_text, doc_status = read_uploaded_file(dokumen)
    progress.progress(40)

    search_text = clean_text(teks) + " " + clean_text(doc_text)
    status.text("Mencocokkan entitas dengan dataset... 75%")
    progress.progress(75)

    best_row, match_status = find_best_row(dataset_df, search_text)
    result = make_result(best_row, teks)

    status.text("Membangun hasil ekstraksi... 100%")
    progress.progress(100)

    st.session_state.last_result = result
    st.session_state.last_dataset_status = dataset_status
    st.session_state.last_doc_status = doc_status
    st.session_state.last_match_status = match_status

    st.success("Proses ekstraksi selesai.")
    return result, doc_status, match_status


# =====================================================
# RENDER OUTPUT
# =====================================================
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


def render_results(result, dataset_status, doc_status, match_status):
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


def render_relation(result):
    st.markdown("## 🔗 Ekstraksi Relasi Bioaktif")
    st.markdown(
        f"""
        <div class="relation-box">
            <p><b>{safe(result["Nama Tanaman"])}</b> → memiliki nama latin → <b>{safe(result["Nama Latin"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → memiliki nama lokal/daerah → <b>{safe(result["Nama Lokal/Daerah"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → menggunakan bagian → <b>{safe(result["Bagian Tanaman"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → mengandung senyawa bioaktif → <b>{safe(result["Zat Bioaktif"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → memiliki khasiat → <b>{safe(result["Khasiat/Efek Terapeutik"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → diolah dengan cara → <b>{safe(result["Cara Pengolahan"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → memiliki dosis/komposisi → <b>{safe(result["Komposisi/Dosis"])}</b></p>
            <p><b>{safe(result["Nama Tanaman"])}</b> → bersumber dari → <b>{safe(result["Sumber Data"])}</b></p>
        </div>
        """,
        unsafe_allow_html=True
    )


def shorten_label(text, max_len=28):
    text = str(text)
    if text.lower() in ["", "nan", "none", "belum terdeteksi"]:
        return "Belum terdeteksi"
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


def wrap_label(text, width=14):
    text = shorten_label(text, 34)
    return "\n".join(textwrap.wrap(text, width=width))


def render_kg(result):
    tanaman = result["Nama Tanaman"] if result["Nama Tanaman"] else "Tanaman"
    latin = result["Nama Latin"] if result["Nama Latin"] else "Nama Latin"
    bagian = result["Bagian Tanaman"] if result["Bagian Tanaman"] else "Bagian Tanaman"
    zat = result["Zat Bioaktif"] if result["Zat Bioaktif"] else "Zat Bioaktif"
    khasiat = result["Khasiat/Efek Terapeutik"] if result["Khasiat/Efek Terapeutik"] else "Khasiat"
    olah = result["Cara Pengolahan"] if result["Cara Pengolahan"] else "Cara Pengolahan"
    dosis = result["Komposisi/Dosis"] if result["Komposisi/Dosis"] else "Dosis"
    sumber = result["Sumber Data"] if result["Sumber Data"] else "Sumber Data"

    st.markdown("## 🕸️ Visualisasi HerbKG 2.0")

    G = nx.DiGraph()

    node_tanaman = "Tanaman\n" + wrap_label(tanaman, 14)
    node_latin = "Nama Latin\n" + wrap_label(latin, 16)
    node_bagian = "Bagian Digunakan\n" + wrap_label(bagian, 14)
    node_zat = "Senyawa Bioaktif\n" + wrap_label(zat, 14)
    node_khasiat = "Khasiat\n" + wrap_label(khasiat, 14)
    node_olah = "Cara Pengolahan\n" + wrap_label(olah, 14)
    node_dosis = "Dosis/Komposisi\n" + wrap_label(dosis, 14)
    node_sumber = "Sumber Data\n" + wrap_label(sumber, 18)

    nodes = {
        node_tanaman: "tanaman",
        node_latin: "latin",
        node_bagian: "bagian",
        node_zat: "senyawa",
        node_khasiat: "khasiat",
        node_olah: "pengolahan",
        node_dosis: "dosis",
        node_sumber: "sumber"
    }

    for node, node_type in nodes.items():
        G.add_node(node, node_type=node_type)

    edges = [
        (node_tanaman, node_latin, "nama latin"),
        (node_tanaman, node_bagian, "bagian digunakan"),
        (node_tanaman, node_zat, "mengandung"),
        (node_zat, node_khasiat, "mendukung khasiat"),
        (node_tanaman, node_khasiat, "memiliki khasiat"),
        (node_tanaman, node_olah, "cara pengolahan"),
        (node_tanaman, node_dosis, "dosis/komposisi"),
        (node_tanaman, node_sumber, "sumber data")
    ]

    for source, target, relation in edges:
        G.add_edge(source, target, label=relation)

    pos = {
        node_tanaman: (0, 0),
        node_latin: (-2.5, 1.6),
        node_bagian: (2.5, 1.6),
        node_zat: (-2.7, 0),
        node_khasiat: (0, -1.8),
        node_olah: (2.7, 0),
        node_dosis: (-2.2, -1.8),
        node_sumber: (2.2, -1.8),
    }

    node_colors = []
    node_edge_colors = []
    node_sizes = []

    for node in G.nodes():
        node_type = G.nodes[node]["node_type"]

        if node_type == "tanaman":
            node_colors.append("#ff5a5f")
            node_edge_colors.append("#ff2028")
            node_sizes.append(5400)
        elif node_type == "senyawa":
            node_colors.append("#ffd6a5")
            node_edge_colors.append("#f97316")
            node_sizes.append(5200)
        elif node_type in ["latin", "khasiat", "sumber"]:
            node_colors.append("#f3c4fb")
            node_edge_colors.append("#a855f7")
            node_sizes.append(5300)
        else:
            node_colors.append("#86efac")
            node_edge_colors.append("#16a34a")
            node_sizes.append(5000)

    fig, ax = plt.subplots(figsize=(13, 8))
    fig.patch.set_facecolor("#fbf7ff")
    ax.set_facecolor("#fbf7ff")

    nx.draw_networkx_edges(
        G,
        pos,
        ax=ax,
        arrows=True,
        arrowsize=24,
        arrowstyle="-|>",
        width=2.4,
        edge_color="#9ca3af",
        connectionstyle="arc3,rad=0.08"
    )

    nx.draw_networkx_nodes(
        G,
        pos,
        ax=ax,
        node_color=node_colors,
        edgecolors=node_edge_colors,
        linewidths=3.5,
        node_size=node_sizes
    )

    nx.draw_networkx_labels(
        G,
        pos,
        ax=ax,
        font_size=9.3,
        font_weight="bold",
        font_color="#111111"
    )

    edge_labels = nx.get_edge_attributes(G, "label")

    nx.draw_networkx_edge_labels(
        G,
        pos,
        edge_labels=edge_labels,
        ax=ax,
        font_size=8.3,
        font_color="#111111",
        rotate=False,
        label_pos=0.55,
        bbox=dict(
            boxstyle="round,pad=0.25",
            fc="#ffffff",
            ec="#e5e7eb",
            alpha=0.95
        )
    )

    ax.set_title(
        "HerbKG 2.0: Relasi Tanaman Herbal, Senyawa Bioaktif, Khasiat, dan Sumber Data",
        fontsize=16,
        fontweight="bold",
        color="#0b7a45",
        pad=18
    )

    ax.set_xlim(-3.4, 3.4)
    ax.set_ylim(-2.6, 2.3)
    ax.axis("off")
    plt.tight_layout()

    st.pyplot(fig, use_container_width=True)


def render_downstream():
    st.markdown("""
    <div class="white-card">
        <h2 style="color:#064e3b;">📦 Aplikasi Downstream</h2>
        <p>Bagian ini menampilkan pemanfaatan hasil ekstraksi dan HerbKG 2.0 untuk analisis lanjutan.</p>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)

    with c1:
        st.markdown("""
        <div class="down-card">
            <h3>1. Analisis Deskriptif</h3>
            <p>Ringkasan statistik tanaman, senyawa, relasi, dan sumber data.</p>
            <table class="mini-table">
                <tr><th>Tanaman</th><th>Relasi</th></tr>
                <tr><td>Kelor</td><td>121</td></tr>
                <tr><td>Jahe</td><td>98</td></tr>
                <tr><td>Sirih</td><td>86</td></tr>
                <tr><td>Kayu Manis</td><td>75</td></tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

    with c2:
        st.markdown("""
        <div class="down-card">
            <h3>2. Query Graf Berbasis Bukti</h3>
            <p>Menelusuri bukti relasi tanaman, senyawa, khasiat, dan sumber literatur.</p>
            <table class="mini-table">
                <tr><th>Entitas</th><th>Relasi</th><th>Bukti</th></tr>
                <tr><td>Tanaman</td><td>mengandung</td><td>Senyawa</td></tr>
                <tr><td>Senyawa</td><td>mendukung</td><td>Khasiat</td></tr>
                <tr><td>Khasiat</td><td>bersumber</td><td>Literatur</td></tr>
            </table>
        </div>
        """, unsafe_allow_html=True)

    with c3:
        st.markdown("""
        <div class="down-card">
            <h3>3. Analisis Kemiripan</h3>
            <p>Menemukan tanaman herbal yang mirip berdasarkan senyawa dan khasiat.</p>
            <div class="mini-flow">Kelor → 0,56</div>
            <div class="mini-flow">Sirih → 0,41</div>
            <div class="mini-flow">Jahe → 0,39</div>
            <div class="mini-flow">Kunyit → 0,31</div>
            <div class="mini-flow">Kayu Manis → 0,28</div>
        </div>
        """, unsafe_allow_html=True)

    with c4:
        st.markdown("""
        <div class="down-card">
            <h3>4. Repurposing Obat Herbal</h3>
            <p>Prediksi pemanfaatan ulang tanaman herbal berdasarkan relasi graf.</p>
            <div class="mini-flow">Penyakit → Target Dikenal</div>
            <div class="mini-flow">Target → Tanaman Terkait</div>
            <div class="mini-flow">Prediksi Kandidat Herbal</div>
            <div class="mini-flow">Peringkat Rekomendasi</div>
        </div>
        """, unsafe_allow_html=True)


# =====================================================
# SIDEBAR
# =====================================================
st.sidebar.markdown("# 🌿 HyTBIONEX")
st.sidebar.markdown("Ekstraksi Informasi Bioaktif & HerbKG 2.0")
st.sidebar.markdown("---")
st.sidebar.markdown("### ANALISIS DATA")

menu = st.sidebar.radio(
    "Pilih Menu",
    [
        "🏠 Dashboard",
        "🌿 Input Tanaman",
        "📁 Upload Dokumen",
        "📋 Hasil Ekstraksi Entitas",
        "🔗 Relation Extraction",
        "🕸️ HerbKG 2.0 Explorer",
        "📦 Aplikasi Downstream",
        "📊 Statistik & Analitik",
        "⚙️ Pengaturan",
        "ℹ️ Tentang Aplikasi"
    ],
    label_visibility="collapsed"
)

st.sidebar.markdown("---")
st.sidebar.markdown("""
<div class="status-pill">
    ✅ Sistem Aktif<br>
    <small>Semua layanan berjalan normal</small>
</div>
""", unsafe_allow_html=True)

# =====================================================
# HEADER
# =====================================================
st.markdown("""
<div class="top-header">
    <div>
        <h1>🌿 HyTBIONEX</h1>
        <p>Hybrid Transformer Pipeline</p>
    </div>
    <div>
        <h2 style="margin:0;color:#bbf7d0;">Analisis Bioaktif & HerbKG 2.0</h2>
        <p>Bioactive Information Extraction & Enhanced Herb Knowledge Graph</p>
    </div>
    <div class="status-pill">🟢 Model Status<br>Aktif</div>
</div>
""", unsafe_allow_html=True)

dataset_df, dataset_status = load_dataset()
file_widget_key = f"uploaded_doc_{st.session_state.file_key_suffix}"
uploaded_active = st.session_state.get(file_widget_key, None) is not None

# =====================================================
# DASHBOARD
# =====================================================
if menu == "🏠 Dashboard":
    left, right = st.columns([2.2, 1])

    with left:
        st.markdown("""
        <div class="hero">
            <h1>HyTBIONEX</h1>
            <h2>Platform Cerdas Ekstraksi Informasi Bioaktif</h2>
            <p>
            Mengintegrasikan pemrosesan bahasa alami, ekstraksi entitas,
            ekstraksi relasi, dan konstruksi HerbKG 2.0 untuk menemukan
            potensi bioaktif tanaman herbal Indonesia.
            </p>
        </div>
        """, unsafe_allow_html=True)

    with right:
        st.markdown("""
        <div class="preview-card">
            <h3 style="color:#064e3b;">Pratinjau HerbKG 2.0</h3>
            <p style="text-align:center;font-size:18px;">
            🌿 <b>Tanaman</b><br>
            ↙️ 🔬 Nama Latin &nbsp;&nbsp; 🍃 Bagian<br>
            ↘️ 🧪 Senyawa &nbsp;&nbsp; 💚 Khasiat
            </p>
        </div>
        """, unsafe_allow_html=True)

    m1, m2, m3, m4 = st.columns(4)
    m1.markdown(f"<div class='metric-card'><h4>🌱 Total Data</h4><h2>{len(dataset_df)}</h2></div>", unsafe_allow_html=True)
    m2.markdown("<div class='metric-card'><h4>🧪 Total Senyawa</h4><h2>Dataset</h2></div>", unsafe_allow_html=True)
    m3.markdown("<div class='metric-card'><h4>💚 Total Khasiat</h4><h2>Dataset</h2></div>", unsafe_allow_html=True)
    m4.markdown("<div class='metric-card'><h4>🔗 Relasi</h4><h2>Triplet</h2></div>", unsafe_allow_html=True)

    st.markdown(
        f"""
        <div class="lilac-card">
            <h3>📌 Status Dataset</h3>
            <p>{safe(dataset_status)}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    render_downstream()

# =====================================================
# INPUT / UPLOAD
# =====================================================
elif menu in ["🌿 Input Tanaman", "📁 Upload Dokumen"]:
    col1, col2 = st.columns(2)

    with col1:
        h1, b1 = st.columns([12, 1])
        with h1:
            st.markdown("""
            <div class="green-card">
                <h2>🌿 1. Input Kata / Kalimat</h2>
                <p>Masukkan nama lokal tanaman atau kalimat deskriptif.</p>
            </div>
            """, unsafe_allow_html=True)
        with b1:
            st.markdown('<div class="small-x">', unsafe_allow_html=True)
            st.button("×", key="clear_input", on_click=clear_input_text)
            st.markdown('</div>', unsafe_allow_html=True)

        if uploaded_active:
            teks = ""
        else:
            teks = st.text_area(
                "Input Data Tanaman",
                placeholder="Contoh: Jahe, Kunyit, Sambiloto, Kayu Manis...",
                height=180,
                key="input_text_area"
            )

    with col2:
        h2, b2 = st.columns([12, 1])
        with h2:
            st.markdown("""
            <div class="green-card">
                <h2>📁 2. Upload Dokumen</h2>
                <p>Unggah dokumen untuk diekstraksi informasinya.</p>
            </div>
            """, unsafe_allow_html=True)
        with b2:
            st.markdown('<div class="small-x">', unsafe_allow_html=True)
            st.button("×", key="clear_file", on_click=clear_uploaded_file)
            st.markdown('</div>', unsafe_allow_html=True)

        dokumen = st.file_uploader(
            "Upload Dokumen",
            type=["pdf", "txt", "csv", "xlsx", "xls"],
            key=file_widget_key,
            on_change=clear_text_when_upload
        )

    if st.button("🔎 Proses Analisis", use_container_width=True):
        result, doc_status, match_status = run_extraction(teks, dokumen, dataset_df, dataset_status)
        render_results(result, dataset_status, doc_status, match_status)
        render_relation(result)
        render_kg(result)

# =====================================================
# HASIL
# =====================================================
elif menu == "📋 Hasil Ekstraksi Entitas":
    if st.session_state.last_result:
        render_results(
            st.session_state.last_result,
            st.session_state.last_dataset_status,
            st.session_state.last_doc_status,
            st.session_state.last_match_status
        )
    else:
        st.warning("Belum ada hasil ekstraksi. Buka menu Input Tanaman atau Upload Dokumen, lalu klik Proses Analisis.")

elif menu == "🔗 Relation Extraction":
    if st.session_state.last_result:
        render_relation(st.session_state.last_result)
    else:
        st.warning("Belum ada hasil relasi. Jalankan proses analisis terlebih dahulu.")

elif menu == "🕸️ HerbKG 2.0 Explorer":
    if st.session_state.last_result:
        render_kg(st.session_state.last_result)
    else:
        st.warning("Belum ada HerbKG. Jalankan proses analisis terlebih dahulu.")

elif menu == "📦 Aplikasi Downstream":
    render_downstream()

elif menu == "📊 Statistik & Analitik":
    st.markdown("""
    <div class="white-card">
        <h2 style="color:#064e3b;">📊 Statistik & Analitik Dataset</h2>
        <p>Menampilkan cuplikan dataset herbal yang digunakan dalam sistem.</p>
    </div>
    """, unsafe_allow_html=True)
    st.dataframe(dataset_df.head(30), use_container_width=True)

elif menu == "⚙️ Pengaturan":
    st.markdown("""
    <div class="lilac-card">
        <h2>⚙️ Pengaturan Sistem</h2>
        <p>Menu ini dapat digunakan untuk konfigurasi model, dataset, dan tampilan sistem.</p>
    </div>
    """, unsafe_allow_html=True)

elif menu == "ℹ️ Tentang Aplikasi":
    st.markdown("""
    <div class="lilac-card">
        <h2>ℹ️ Tentang HyTBIONEX</h2>
        <p>
        HyTBIONEX adalah prototipe sistem untuk ekstraksi informasi bioaktif tanaman herbal Indonesia,
        ekstraksi relasi, dan pembangunan Enhanced Herb Knowledge Graph 2.0.
        </p>
    </div>
    """, unsafe_allow_html=True)
