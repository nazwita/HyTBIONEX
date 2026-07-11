import os
import re
import html
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go


# =========================================================
# KONFIGURASI APLIKASI
# =========================================================
APP_TITLE = "HyTBIONEX"
DATASET_FILE = "Data set 20098+ Gambar.xlsx"
IMAGE_TABLE_FILE = "Gambar tanaman herbal.xlsx"
ASSET_DIR = "assets"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded"
)


# =========================================================
# SESSION STATE
# =========================================================
if "page" not in st.session_state:
    st.session_state.page = "🏠 Dashboard"

if "last_result" not in st.session_state:
    st.session_state.last_result = None

if "last_image" not in st.session_state:
    st.session_state.last_image = None

if "last_status" not in st.session_state:
    st.session_state.last_status = {}

if "downstream_view" not in st.session_state:
    st.session_state.downstream_view = "awal"


# =========================================================
# FUNGSI DASAR
# =========================================================
def html_block(code):
    st.markdown(code, unsafe_allow_html=True)


def safe_text(x):
    if x is None:
        return ""
    return html.escape(str(x))


def clean_text(x):
    x = str(x).lower()
    x = re.sub(r"[^a-zA-Z0-9À-ÿ\s]", " ", x)
    x = re.sub(r"\s+", " ", x)
    return x.strip()


def slugify_filename(text):
    text = clean_text(text)
    text = text.replace(" ", "_")
    text = re.sub(r"_+", "_", text)
    return text.strip("_")


def normalize_colname(col):
    col = str(col).strip().lower()
    col = col.replace("_", " ")
    col = col.replace("/", " ")
    col = re.sub(r"[^a-zA-Z0-9À-ÿ\s]", " ", col)
    col = re.sub(r"\s+", " ", col)
    return col.strip()


def find_col(df, candidates):
    if df is None or df.empty:
        return None

    normalized_cols = {normalize_colname(c): c for c in df.columns}

    for cand in candidates:
        cand_norm = normalize_colname(cand)
        if cand_norm in normalized_cols:
            return normalized_cols[cand_norm]

    for cand in candidates:
        cand_norm = normalize_colname(cand)
        for norm_col, original_col in normalized_cols.items():
            if cand_norm in norm_col or norm_col in cand_norm:
                return original_col

    return None


def value_from_row(row, col):
    if row is None or col is None:
        return "Belum terdeteksi"

    try:
        value = row.get(col, "")
    except Exception:
        return "Belum terdeteksi"

    if pd.isna(value) or str(value).strip() == "":
        return "Belum terdeteksi"

    return str(value).strip()


def safe_dataframe(df, cols=None, n=50):
    if df is None or df.empty:
        st.info("Data belum tersedia.")
        return

    try:
        if cols:
            cols = [c for c in cols if c is not None and c in df.columns]
            if cols:
                st.dataframe(df[cols].drop_duplicates().head(n))
            else:
                st.dataframe(df.head(n))
        else:
            st.dataframe(df.head(n))
    except Exception as e:
        st.warning("Tabel belum dapat ditampilkan.")
        st.code(str(e))


# =========================================================
# LOAD DATASET
# =========================================================
@st.cache_data(show_spinner=False)
def load_dataset():
    dataset_path = ""

    if os.path.exists(DATASET_FILE):
        dataset_path = DATASET_FILE
    else:
        excel_files = [
            f for f in os.listdir(".")
            if f.lower().endswith((".xlsx", ".xls"))
            and f != IMAGE_TABLE_FILE
        ]

        if excel_files:
            dataset_path = excel_files[0]

    if not dataset_path:
        return pd.DataFrame(), "Dataset Excel utama belum ditemukan."

    try:
        sheets = pd.read_excel(dataset_path, sheet_name=None)
        frames = []

        for sheet_name, df in sheets.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                df = df.fillna("")
                df["__sheet_name__"] = sheet_name
                frames.append(df)

        if not frames:
            return pd.DataFrame(), f"Dataset kosong: {dataset_path}"

        data = pd.concat(frames, ignore_index=True).fillna("")
        data.columns = [str(c).strip() for c in data.columns]

        return data, f"Dataset terbaca: {dataset_path} | Total data: {len(data)} baris"

    except Exception as e:
        return pd.DataFrame(), f"Gagal membaca dataset: {e}"


@st.cache_data(show_spinner=False)
def load_image_mapping():
    """
    Membaca file Gambar tanaman herbal.xlsx jika ada.
    Kolom Gambar harus berisi teks path, contoh:
    assets/serai.jpg
    assets/kayu_manis.jpg
    """
    mapping = {}

    if not os.path.exists(IMAGE_TABLE_FILE):
        return mapping

    try:
        sheets = pd.read_excel(IMAGE_TABLE_FILE, sheet_name=None)

        image_df = pd.DataFrame()
        for _, df in sheets.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                image_df = df.fillna("")
                break

        if image_df.empty:
            return mapping

        image_df.columns = [str(c).strip() for c in image_df.columns]

        col_nama = find_col(image_df, ["Nama Tanaman", "Tanaman", "Nama"])
        col_latin = find_col(image_df, ["Nama Latin", "Latin"])
        col_gambar = find_col(image_df, ["Gambar", "Image", "Foto", "File Gambar", "Path Gambar"])

        for _, row in image_df.iterrows():
            nama = value_from_row(row, col_nama)
            latin = value_from_row(row, col_latin)
            gambar = value_from_row(row, col_gambar)

            if gambar == "Belum terdeteksi":
                gambar = ""

            for key in [clean_text(nama), clean_text(latin)]:
                if key:
                    mapping[key] = {
                        "nama": nama,
                        "latin": latin,
                        "gambar": gambar
                    }

    except Exception:
        pass

    return mapping


# =========================================================
# KOLOM DATASET
# =========================================================
def get_columns(df):
    return {
        "nama": find_col(df, ["Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"]),
        "latin": find_col(df, ["Nama Latin", "Nama_Latin", "Latin"]),
        "lokal": find_col(df, ["Nama Lokal/Daerah", "Nama Lokal", "Nama Daerah", "Bahasa Daerah", "Bahasa_Daerah"]),
        "bagian": find_col(df, ["Bagian Tanaman", "Bagian Digunakan", "Bagian_Digunakan", "Bagian"]),
        "senyawa": find_col(df, ["Zat Bioaktif", "Senyawa Bioaktif", "Senyawa_Bioaktif", "Compound", "Senyawa", "Kandungan", "Kandungan Kimia"]),
        "khasiat": find_col(df, ["Khasiat/Efek Terapeutik", "Khasiat", "Manfaat", "Benefit", "Biological Activity", "Biological_Activity"]),
        "pengolahan": find_col(df, ["Cara Pengolahan", "Cara_Pengolahan", "Pengolahan", "Cara Pemakaian"]),
        "dosis": find_col(df, ["Komposisi/Dosis", "Komposisi /Dosis", "Dosis", "Komposisi"]),
        "sumber": find_col(df, ["Sumber Data", "Sumber_Data", "Sumber", "Referensi"]),
        "gambar": find_col(df, ["Gambar", "Image", "Foto", "File Gambar", "Path Gambar", "Nama File Gambar"]),
    }


# =========================================================
# BACA DOKUMEN UPLOAD
# =========================================================
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
                return "", f"PDF terbaca, tetapi teks kosong: {uploaded_file.name}. Kemungkinan PDF berupa scan/gambar."

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


# =========================================================
# MATCHING DATASET
# =========================================================
def score_row(row, search_text, cols):
    score = 0

    nama = clean_text(value_from_row(row, cols["nama"]))
    latin = clean_text(value_from_row(row, cols["latin"]))
    lokal = clean_text(value_from_row(row, cols["lokal"]))

    if nama and nama != "belum terdeteksi":
        if nama in search_text:
            score += 200
        for token in nama.split():
            if len(token) >= 3 and token in search_text:
                score += 30

    if latin and latin != "belum terdeteksi":
        if latin in search_text:
            score += 170
        for token in latin.split():
            if len(token) >= 4 and token in search_text:
                score += 25

    if lokal and lokal != "belum terdeteksi":
        parts = re.split(r"[,;/|]", lokal)
        for p in parts:
            p = clean_text(p)
            if p and len(p) >= 3 and p in search_text:
                score += 90

    return score


def find_best_match(df, search_text):
    if df is None or df.empty:
        return None, "Dataset belum terbaca."

    search_text = clean_text(search_text)

    if not search_text:
        return None, "Input tanaman dan dokumen masih kosong."

    cols = get_columns(df)

    best_row = None
    best_score = 0

    try:
        for _, row in df.iterrows():
            score = score_row(row, search_text, cols)

            if score > best_score:
                best_score = score
                best_row = row
    except Exception as e:
        return None, f"Gagal mencocokkan data: {e}"

    if best_row is not None and best_score > 0:
        nama = value_from_row(best_row, cols["nama"])
        latin = value_from_row(best_row, cols["latin"])
        return best_row, f"Entitas cocok dengan dataset: {nama} / {latin} | Skor: {best_score}"

    return None, "Tidak ditemukan kecocokan entitas tanaman pada dataset."


def extract_result(row, input_text, df):
    cols = get_columns(df)

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
            "Sumber Data": "Belum terdeteksi",
            "Gambar": "Belum terdeteksi",
        }

    return {
        "Nama Tanaman": value_from_row(row, cols["nama"]),
        "Nama Latin": value_from_row(row, cols["latin"]),
        "Nama Lokal/Daerah": value_from_row(row, cols["lokal"]),
        "Bagian Tanaman": value_from_row(row, cols["bagian"]),
        "Zat Bioaktif": value_from_row(row, cols["senyawa"]),
        "Khasiat/Efek Terapeutik": value_from_row(row, cols["khasiat"]),
        "Cara Pengolahan": value_from_row(row, cols["pengolahan"]),
        "Komposisi/Dosis": value_from_row(row, cols["dosis"]),
        "Sumber Data": value_from_row(row, cols["sumber"]),
        "Gambar": value_from_row(row, cols["gambar"]),
    }


# =========================================================
# GAMBAR TANAMAN
# =========================================================
def path_if_exists(path):
    try:
        if path and os.path.exists(path):
            return path
    except Exception:
        pass
    return None


def scan_assets_by_name(nama, latin):
    if not os.path.isdir(ASSET_DIR):
        return None

    keys = []
    for item in [nama, latin]:
        slug = slugify_filename(item)
        if slug:
            keys.append(slug)

    image_exts = [".jpg", ".jpeg", ".png", ".webp"]

    try:
        for file_name in os.listdir(ASSET_DIR):
            file_path = os.path.join(ASSET_DIR, file_name)

            if not os.path.isfile(file_path):
                continue

            stem = slugify_filename(Path(file_name).stem)
            ext = Path(file_name).suffix.lower()

            if ext not in image_exts:
                continue

            for key in keys:
                if key == stem or key in stem or stem in key:
                    return file_path
    except Exception:
        return None

    return None


def find_plant_image(result, image_mapping):
    try:
        gambar = result.get("Gambar", "")
        nama = result.get("Nama Tanaman", "")
        latin = result.get("Nama Latin", "")
    except Exception:
        return None

    candidates = []

    # Dari kolom Gambar di dataset utama
    if gambar and gambar != "Belum terdeteksi":
        candidates.extend([
            gambar,
            os.path.join(ASSET_DIR, gambar),
            os.path.join("gambar", gambar),
            os.path.join("images", gambar),
        ])

    # Dari file Gambar tanaman herbal.xlsx
    for key in [clean_text(nama), clean_text(latin)]:
        try:
            if key in image_mapping:
                map_gambar = image_mapping[key].get("gambar", "")

                if map_gambar:
                    candidates.extend([
                        map_gambar,
                        os.path.join(ASSET_DIR, map_gambar),
                        os.path.join("gambar", map_gambar),
                        os.path.join("images", map_gambar),
                    ])
        except Exception:
            pass

    # Otomatis dari nama tanaman dan latin
    for item in [nama, latin]:
        slug = slugify_filename(item)

        if slug:
            for ext in ["jpg", "jpeg", "png", "webp"]:
                candidates.extend([
                    os.path.join(ASSET_DIR, f"{slug}.{ext}"),
                    os.path.join("gambar", f"{slug}.{ext}"),
                    os.path.join("images", f"{slug}.{ext}"),
                    f"{slug}.{ext}",
                ])

    for path in candidates:
        found = path_if_exists(path)
        if found:
            return found

    return scan_assets_by_name(nama, latin)


# =========================================================
# PROSES EKSTRAKSI AMAN
# =========================================================
def run_extraction(text_input, uploaded_file, df, dataset_status, image_mapping):
    doc_text = ""
    doc_status = "Tidak ada dokumen yang diupload."
    match_status = "Belum diproses."

    try:
        doc_text, doc_status = read_uploaded_file(uploaded_file)
    except Exception as e:
        doc_status = f"Dokumen gagal dibaca: {e}"

    combined_text = f"{text_input} {doc_text}"

    try:
        row, match_status = find_best_match(df, combined_text)
    except Exception as e:
        row = None
        match_status = f"Proses pencocokan gagal: {e}"

    try:
        result = extract_result(row, text_input, df)
    except Exception as e:
        result = {
            "Nama Tanaman": text_input if text_input else "Belum terdeteksi",
            "Nama Latin": "Belum terdeteksi",
            "Nama Lokal/Daerah": "Belum terdeteksi",
            "Bagian Tanaman": "Belum terdeteksi",
            "Zat Bioaktif": "Belum terdeteksi",
            "Khasiat/Efek Terapeutik": "Belum terdeteksi",
            "Cara Pengolahan": "Belum terdeteksi",
            "Komposisi/Dosis": "Belum terdeteksi",
            "Sumber Data": "Belum terdeteksi",
            "Gambar": "Belum terdeteksi",
        }
        match_status = f"Ekstraksi gagal: {e}"

    try:
        image_path = find_plant_image(result, image_mapping)
    except Exception as e:
        image_path = None
        match_status = f"{match_status} | Gambar gagal dicari: {e}"

    st.session_state.last_result = result
    st.session_state.last_image = image_path
    st.session_state.last_status = {
        "dataset": dataset_status,
        "document": doc_status,
        "match": match_status,
    }

    return result, image_path, doc_status, match_status


# =========================================================
# GRAFIK DAN KG
# =========================================================
def make_descriptive_chart(df, column=None, title="Analisis Deskriptif"):
    if df is None or df.empty:
        return None

    if column is None:
        cols = get_columns(df)
        column = cols["nama"]

    if column is None or column not in df.columns:
        return None

    try:
        chart_df = (
            df[column]
            .astype(str)
            .str.strip()
            .replace("", pd.NA)
            .dropna()
            .value_counts()
            .head(10)
            .reset_index()
        )

        chart_df.columns = ["Kategori", "Jumlah Data"]

        fig = px.bar(
            chart_df,
            x="Jumlah Data",
            y="Kategori",
            orientation="h",
            text="Jumlah Data",
            color="Jumlah Data",
            color_continuous_scale=["#bbf7d0", "#22c55e", "#047857"],
            title=title
        )

        fig.update_layout(
            height=430,
            plot_bgcolor="white",
            paper_bgcolor="white",
            title_font=dict(size=20, color="#064e3b"),
            xaxis_title="Jumlah Data",
            yaxis_title="Kategori",
            margin=dict(l=20, r=20, t=60, b=20),
        )

        fig.update_traces(textposition="outside")
        fig.update_yaxes(autorange="reversed")

        return fig

    except Exception:
        return None


def short_label(text, max_len=24):
    text = str(text)

    if text in ["", "nan", "None", "Belum terdeteksi"]:
        return "Belum terdeteksi"

    if len(text) <= max_len:
        return text

    return text[:max_len] + "..."


def make_kg_graph(result):
    try:
        tanaman = short_label(result["Nama Tanaman"], 18)
        latin = short_label(result["Nama Latin"], 24)
        bagian = short_label(result["Bagian Tanaman"], 18)
        senyawa = short_label(result["Zat Bioaktif"], 24)
        khasiat = short_label(result["Khasiat/Efek Terapeutik"], 22)
        pengolahan = short_label(result["Cara Pengolahan"], 22)
        dosis = short_label(result["Komposisi/Dosis"], 22)
        sumber = short_label(result["Sumber Data"], 26)
    except Exception:
        tanaman = "Tanaman"
        latin = "Nama Latin"
        bagian = "Bagian"
        senyawa = "Senyawa"
        khasiat = "Khasiat"
        pengolahan = "Pengolahan"
        dosis = "Dosis"
        sumber = "Sumber"

    nodes = [
        {"id": "tanaman", "label": tanaman, "x": 0, "y": 0, "color": "#047857", "size": 76},
        {"id": "latin", "label": latin, "x": -2.4, "y": 1.45, "color": "#e9d5ff", "size": 58},
        {"id": "bagian", "label": bagian, "x": 2.4, "y": 1.45, "color": "#bbf7d0", "size": 58},
        {"id": "senyawa", "label": senyawa, "x": -2.5, "y": -0.05, "color": "#fed7aa", "size": 62},
        {"id": "khasiat", "label": khasiat, "x": 0, "y": -1.65, "color": "#fbcfe8", "size": 62},
        {"id": "pengolahan", "label": pengolahan, "x": 2.55, "y": -0.1, "color": "#bbf7d0", "size": 58},
        {"id": "dosis", "label": dosis, "x": -2.2, "y": -1.75, "color": "#fed7aa", "size": 58},
        {"id": "sumber", "label": sumber, "x": 2.2, "y": -1.75, "color": "#e9d5ff", "size": 58},
    ]

    node_map = {n["id"]: n for n in nodes}

    edges = [
        ("tanaman", "latin", "nama latin"),
        ("tanaman", "bagian", "bagian digunakan"),
        ("tanaman", "senyawa", "mengandung"),
        ("senyawa", "khasiat", "mendukung khasiat"),
        ("tanaman", "khasiat", "memiliki khasiat"),
        ("tanaman", "pengolahan", "cara pengolahan"),
        ("tanaman", "dosis", "dosis/komposisi"),
        ("tanaman", "sumber", "sumber data"),
    ]

    fig = go.Figure()

    for source, target, label in edges:
        a = node_map[source]
        b = node_map[target]

        fig.add_trace(go.Scatter(
            x=[a["x"], b["x"]],
            y=[a["y"], b["y"]],
            mode="lines",
            line=dict(color="#94a3b8", width=2.2),
            hoverinfo="none",
            showlegend=False
        ))

        fig.add_annotation(
            x=(a["x"] + b["x"]) / 2,
            y=(a["y"] + b["y"]) / 2,
            text=label,
            showarrow=False,
            font=dict(size=12, color="#111111"),
            bgcolor="rgba(255,255,255,0.82)",
            bordercolor="rgba(226,232,240,0.9)",
            borderwidth=1,
            borderpad=3
        )

    for node in nodes:
        border_color = "#047857"

        if node["id"] in ["senyawa", "dosis"]:
            border_color = "#f97316"
        elif node["id"] in ["latin", "sumber"]:
            border_color = "#a855f7"
        elif node["id"] == "khasiat":
            border_color = "#ec4899"

        fig.add_trace(go.Scatter(
            x=[node["x"]],
            y=[node["y"]],
            mode="markers+text",
            marker=dict(
                size=node["size"],
                color=node["color"],
                line=dict(color=border_color, width=4)
            ),
            text=[node["label"]],
            textposition="middle center",
            textfont=dict(size=14, color="#111111", family="Arial Black"),
            hoverinfo="text",
            showlegend=False
        ))

    fig.update_layout(
        height=620,
        plot_bgcolor="#fbf7ff",
        paper_bgcolor="#fbf7ff",
        margin=dict(l=10, r=10, t=20, b=20),
        xaxis=dict(visible=False, range=[-3.2, 3.2]),
        yaxis=dict(visible=False, range=[-2.4, 2.2]),
    )

    return fig


# =========================================================
# CSS TAMPILAN
# =========================================================
html_block("""
<style>
.stApp {
    background: #f5fbf7 !important;
}

.main .block-container {
    padding-top: 1.2rem;
    max-width: 1500px;
}

/* SIDEBAR */
[data-testid="stSidebar"] {
    background:
        radial-gradient(circle at bottom left, rgba(34,197,94,0.25), transparent 28%),
        radial-gradient(circle at top right, rgba(16,185,129,0.18), transparent 35%),
        linear-gradient(180deg, #021f16 0%, #043b2c 45%, #065f46 100%) !important;
    border-right: 1px solid rgba(255,255,255,0.12);
}

[data-testid="stSidebar"] * {
    color: #f8fff8 !important;
}

[data-testid="stSidebar"] .stButton > button {
    background: transparent !important;
    color: #f8fff8 !important;
    border: none !important;
    box-shadow: none !important;
    text-align: left !important;
    justify-content: flex-start !important;
    padding: 12px 14px !important;
    border-radius: 14px !important;
    font-size: 16px !important;
    font-weight: 650 !important;
    margin-bottom: 5px !important;
    width: 100% !important;
}

[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.12) !important;
    color: #ffffff !important;
    transform: translateX(3px);
}

.nav-active {
    background:
        linear-gradient(90deg, rgba(34,197,94,0.38), rgba(16,185,129,0.20)) !important;
    border: 1px solid rgba(134,239,172,0.28);
    box-shadow: 0 0 22px rgba(34,197,94,0.22);
    border-radius: 15px;
    padding: 13px 15px;
    margin-bottom: 8px;
    color: #ffffff !important;
    font-size: 16px;
    font-weight: 800;
}

.sidebar-section-title {
    color: #86efac !important;
    font-size: 13px !important;
    font-weight: 900 !important;
    letter-spacing: 0.6px;
    margin-top: 22px;
    margin-bottom: 10px;
    text-transform: uppercase;
}

.sidebar-line {
    height: 1px;
    background: rgba(255,255,255,0.14);
    margin: 18px 0;
}

.sidebar-footer {
    margin-top: 28px;
    padding: 18px;
    border-radius: 18px;
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.14);
}

.sidebar-footer h3 {
    color: #ffffff !important;
    margin: 0 0 6px 0 !important;
    font-size: 17px !important;
    font-weight: 900 !important;
}

.sidebar-footer p {
    color: #d1fae5 !important;
    margin: 0 !important;
    font-size: 12px !important;
}

/* HEADER */
.top-header {
    background:
        linear-gradient(135deg, rgba(1,50,32,0.98), rgba(6,78,59,0.98), rgba(5,95,70,0.98));
    padding: 24px 30px;
    border-radius: 0 0 28px 28px;
    color: white;
    margin-bottom: 22px;
    box-shadow: 0 16px 38px rgba(0,0,0,0.18);
    display: grid;
    grid-template-columns: 1.2fr 1.5fr 0.85fr;
    gap: 18px;
    align-items: center;
}

.top-header h1 {
    margin: 0;
    font-size: 38px;
    font-weight: 900;
    color: #ffffff;
}

.top-header h2 {
    margin: 0;
    font-size: 28px;
    font-weight: 900;
    color: #bbf7d0;
}

.top-header p {
    margin: 5px 0 0 0;
    color: #ecfdf5;
}

.top-user {
    display: flex;
    gap: 12px;
    justify-content: flex-end;
    align-items: center;
}

.top-pill {
    background: rgba(255,255,255,0.10);
    padding: 11px 14px;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.12);
    color: #ffffff;
    font-weight: 800;
    font-size: 14px;
}

.user-card {
    background: rgba(255,255,255,0.10);
    padding: 10px 14px;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.12);
    color: #ffffff;
    font-weight: 800;
    min-width: 145px;
}

/* KONTEN */
.hero-banner {
    background:
        radial-gradient(circle at 70% 40%, rgba(187,247,208,0.92), transparent 28%),
        linear-gradient(135deg, #ffffff 0%, #ecfdf5 55%, #f8fafc 100%);
    padding: 34px 32px;
    border-radius: 22px;
    min-height: 215px;
    box-shadow: inset 0 0 0 1px rgba(6,78,59,0.08);
}

.hero-banner h2 {
    color: #047857;
    font-size: 30px;
    font-weight: 900;
    margin-bottom: 14px;
}

.hero-banner p {
    color: #12372a;
    font-size: 17px;
    line-height: 1.65;
    max-width: 680px;
}

.input-panel {
    background: rgba(255,255,255,0.97);
    padding: 25px;
    border-radius: 22px;
    box-shadow: 0 12px 28px rgba(15,23,42,0.10);
    border: 1px solid rgba(6,78,59,0.10);
    margin-top: 18px;
    margin-bottom: 18px;
}

.input-panel h3 {
    color: #064e3b;
    margin: 0 0 10px 0;
    font-weight: 900;
}

textarea {
    background: #ffffff !important;
    color: #0f172a !important;
    border: 1px solid #d1d5db !important;
    border-radius: 12px !important;
}

[data-testid="stFileUploader"] section {
    background: #fbfffb !important;
    border: 2px dashed #86efac !important;
    border-radius: 16px !important;
    min-height: 125px;
}

[data-testid="stFileUploader"] section * {
    color: #0f172a !important;
}

.stButton > button {
    background: linear-gradient(90deg, #047857, #059669) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 13px !important;
    font-weight: 900 !important;
    padding: 0.72rem 1rem !important;
    box-shadow: 0 8px 18px rgba(4,120,87,0.18);
}

.result-card {
    background: #ffffff;
    border-left: 7px solid #047857;
    border-radius: 15px;
    padding: 16px;
    min-height: 112px;
    box-shadow: 0 8px 18px rgba(15,23,42,0.09);
    margin-bottom: 12px;
}

.result-card h4 {
    color: #064e3b;
    font-weight: 900;
    margin-bottom: 7px;
}

.result-card p {
    color: #0f172a;
    font-size: 16px;
    line-height: 1.5;
}

.section-title {
    color: #064e3b;
    font-size: 27px;
    font-weight: 900;
    margin-top: 22px;
    margin-bottom: 13px;
}

.lilac-card {
    background: #f3e8ff;
    padding: 22px;
    border-radius: 18px;
    border: 2px solid #c084fc;
    color: #111111;
    margin-bottom: 18px;
}

.down-card {
    background: #ffffff;
    border-radius: 20px;
    padding: 18px;
    box-shadow: 0 10px 24px rgba(15,23,42,0.08);
    border: 1px solid rgba(6,78,59,0.12);
    min-height: 225px;
}

.down-card h3 {
    color: #064e3b;
    text-align: center;
    font-weight: 900;
    font-size: 23px;
    margin-bottom: 14px;
}

.down-card p {
    color: #334155;
    font-size: 15px;
    line-height: 1.55;
    min-height: 70px;
}

.down-btn button {
    background: #f0fdf4 !important;
    border: 1px solid #bbf7d0 !important;
    color: #064e3b !important;
    border-radius: 14px !important;
    font-weight: 900 !important;
    margin-bottom: 8px !important;
    box-shadow: none !important;
}

.footer-status {
    background: rgba(255,255,255,0.94);
    border-radius: 16px;
    padding: 17px;
    border: 1px solid rgba(6,78,59,0.12);
    color:#064e3b;
    font-weight:800;
    text-align:center;
}
</style>
""")


# =========================================================
# LOAD DATA GLOBAL
# =========================================================
df_data, dataset_status = load_dataset()
image_mapping = load_image_mapping()
cols_global = get_columns(df_data)

total_data = len(df_data)

if cols_global["senyawa"]:
    total_senyawa = df_data[cols_global["senyawa"]].astype(str).replace("", pd.NA).dropna().nunique()
else:
    total_senyawa = 0

if cols_global["khasiat"]:
    total_khasiat = df_data[cols_global["khasiat"]].astype(str).replace("", pd.NA).dropna().nunique()
else:
    total_khasiat = 0

total_relasi = total_data * 8 if total_data else 0


# =========================================================
# SIDEBAR
# =========================================================
def set_page(page_name):
    st.session_state.page = page_name


def sidebar_button(label, page_name, key):
    if st.session_state.page == page_name:
        st.sidebar.markdown(
            f'<div class="nav-active">{label}</div>',
            unsafe_allow_html=True
        )
    else:
        if st.sidebar.button(label, key=key):
            set_page(page_name)
            st.rerun()


st.sidebar.markdown("# 🌿 HyTBIONEX")
st.sidebar.markdown("Hybrid Transformer Pipeline")
st.sidebar.markdown('<div class="sidebar-line"></div>', unsafe_allow_html=True)

sidebar_button("🏠 Dashboard", "🏠 Dashboard", "nav_dashboard")

st.sidebar.markdown('<div class="sidebar-section-title">ANALISIS DATA</div>', unsafe_allow_html=True)
sidebar_button("🌿 Input Tanaman", "🌿 Input Tanaman", "nav_input")
sidebar_button("📄 Upload Dokumen", "📁 Upload Dokumen", "nav_upload")
sidebar_button("📋 Hasil Isolasi Entitas", "📋 Hasil Isolasi Entitas", "nav_hasil")
sidebar_button("🔗 Relation Extraction", "🔗 Relation Extraction", "nav_relasi")
sidebar_button("🕸️ HerbKG 2.0 Explorer", "🕸️ HerbKG 2.0 Explorer", "nav_kg")

st.sidebar.markdown('<div class="sidebar-line"></div>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="sidebar-section-title">VISUALISASI</div>', unsafe_allow_html=True)
sidebar_button("📦 Aplikasi Downstream", "📦 Aplikasi Downstream", "nav_downstream")
sidebar_button("📊 Statistik & Analitik", "📊 Statistik & Analitik", "nav_statistik")

st.sidebar.markdown('<div class="sidebar-line"></div>', unsafe_allow_html=True)
st.sidebar.markdown('<div class="sidebar-section-title">MODEL & SISTEM</div>', unsafe_allow_html=True)
sidebar_button("🧩 Training Model", "⚙️ Pengaturan", "nav_training")
sidebar_button("⚙️ Pengaturan", "⚙️ Pengaturan", "nav_setting")
sidebar_button("ℹ️ Tentang Aplikasi", "ℹ️ Tentang Aplikasi", "nav_about")

st.sidebar.markdown("""
<div class="sidebar-footer">
    <h3>🌿 HyTBIONEX</h3>
    <p>© 2025 All rights reserved</p>
</div>
""", unsafe_allow_html=True)

menu = st.session_state.page


# =========================================================
# HEADER
# =========================================================
html_block("""
<div class="top-header">
    <div>
        <h1>🌿 HyTBIONEX</h1>
        <p>Hybrid Transformer Pipeline</p>
    </div>
    <div>
        <h2>Analisis Bioaktif & HerbKG 2.0</h2>
        <p>Bioactive Information Isolation & Enhanced Herb Knowledge Graph</p>
    </div>
    <div class="top-user">
        <div class="top-pill">🟢 Model Status<br>Aktif</div>
        <div class="top-pill">🌙 Dark Mode</div>
        <div class="user-card">👩‍🔬 Nazwita<br><small>Researcher</small></div>
    </div>
</div>
""")


# =========================================================
# RENDER KOMPONEN
# =========================================================
def render_summary_card():
    st.subheader("RINGKASAN DATA")

    m1, m2 = st.columns(2)
    m3, m4 = st.columns(2)

    with m1:
        st.metric("🌿 Total Tanaman", f"{total_data:,}")

    with m2:
        st.metric("🧪 Total Senyawa", f"{total_senyawa:,}")

    with m3:
        st.metric("💗 Total Khasiat", f"{total_khasiat:,}")

    with m4:
        st.metric("🔗 Relasi Triplet", f"{total_relasi:,}")

    st.caption("Dataset dan relasi dihitung otomatis dari file Excel.")


def render_input_area():
    html_block("""
    <div class="input-panel">
        <h3>📝 Input Tanaman dan Dokumen</h3>
        <p>Masukkan nama tanaman atau kalimat. Upload dokumen bersifat opsional. Sistem akan mencocokkan informasi dengan dataset herbal.</p>
    </div>
    """)

    col1, col2 = st.columns([1.1, 1])

    with col1:
        text_input = st.text_area(
            "Input Data Tanaman",
            placeholder="Contoh: Serai, Jahe, Kunyit, Sambiloto, Kayu Manis...",
            height=130,
            key="main_input_text"
        )

    with col2:
        uploaded_file = st.file_uploader(
            "Upload Dokumen",
            type=["pdf", "txt", "csv", "xlsx", "xls"],
            key="main_uploaded_file"
        )

    submit = st.button("🔍 Proses Analisis", key="btn_process_all")

    if submit:
        if not text_input and uploaded_file is None:
            st.warning("Masukkan nama tanaman atau upload dokumen terlebih dahulu.")
            return

        with st.spinner("Sedang memproses ekstraksi informasi bioaktif..."):
            result, image_path, doc_status, match_status = run_extraction(
                text_input,
                uploaded_file,
                df_data,
                dataset_status,
                image_mapping
            )

        st.success("Proses analisis selesai.")
        render_all_outputs(result, image_path, dataset_status, doc_status, match_status)


def render_result_cards(result):
    html_block('<div class="section-title">📋 Hasil Ekstraksi Informasi Bioaktif</div>')

    cards = [
        ("🌿 Nama Tanaman", result.get("Nama Tanaman", "Belum terdeteksi")),
        ("🔬 Nama Latin", result.get("Nama Latin", "Belum terdeteksi")),
        ("🇮🇩 Nama Lokal/Daerah", result.get("Nama Lokal/Daerah", "Belum terdeteksi")),
        ("🍃 Bagian Tanaman", result.get("Bagian Tanaman", "Belum terdeteksi")),
        ("🧪 Zat Bioaktif", result.get("Zat Bioaktif", "Belum terdeteksi")),
        ("💚 Khasiat / Efek Terapeutik", result.get("Khasiat/Efek Terapeutik", "Belum terdeteksi")),
        ("☕ Cara Pengolahan", result.get("Cara Pengolahan", "Belum terdeteksi")),
        ("⚖️ Komposisi / Dosis", result.get("Komposisi/Dosis", "Belum terdeteksi")),
        ("📚 Sumber Data", result.get("Sumber Data", "Belum terdeteksi")),
    ]

    rows = [st.columns(3), st.columns(3), st.columns(3)]
    idx = 0

    for row_cols in rows:
        for col in row_cols:
            title, value = cards[idx]
            with col:
                html_block(f"""
                <div class="result-card">
                    <h4>{safe_text(title)}</h4>
                    <p>{safe_text(value)}</p>
                </div>
                """)
            idx += 1


def render_image_section(result, image_path):
    html_block('<div class="section-title">🖼️ Lampiran Gambar Tanaman</div>')

    if image_path and os.path.exists(image_path):
        c1, c2 = st.columns([1, 2])

        with c1:
            st.image(image_path, caption=result.get("Nama Tanaman", "Tanaman"))

        with c2:
            html_block(f"""
            <div class="lilac-card">
                <h3>🌿 Gambar Tanaman Terkoneksi Dataset</h3>
                <p><b>Nama Tanaman:</b> {safe_text(result.get("Nama Tanaman", "Belum terdeteksi"))}</p>
                <p><b>Nama Latin:</b> {safe_text(result.get("Nama Latin", "Belum terdeteksi"))}</p>
                <p><b>Catatan:</b> Gambar tampil dari kolom <b>Gambar</b> di Excel atau file yang cocok di folder <b>assets</b>.</p>
            </div>
            """)
    else:
        st.info("Gambar belum ditemukan. Isi kolom Gambar di Excel dengan contoh: assets/serai.jpg atau simpan file gambar di folder assets.")


def render_relation_table(result):
    html_block('<div class="section-title">🔗 Relation Extraction</div>')

    rel_df = pd.DataFrame([
        [result.get("Nama Tanaman", "Tanaman"), "memiliki nama latin", result.get("Nama Latin", "Belum terdeteksi")],
        [result.get("Nama Tanaman", "Tanaman"), "memiliki nama lokal/daerah", result.get("Nama Lokal/Daerah", "Belum terdeteksi")],
        [result.get("Nama Tanaman", "Tanaman"), "menggunakan bagian tanaman", result.get("Bagian Tanaman", "Belum terdeteksi")],
        [result.get("Nama Tanaman", "Tanaman"), "mengandung senyawa bioaktif", result.get("Zat Bioaktif", "Belum terdeteksi")],
        [result.get("Nama Tanaman", "Tanaman"), "memiliki khasiat", result.get("Khasiat/Efek Terapeutik", "Belum terdeteksi")],
        [result.get("Nama Tanaman", "Tanaman"), "diolah dengan cara", result.get("Cara Pengolahan", "Belum terdeteksi")],
        [result.get("Nama Tanaman", "Tanaman"), "memiliki dosis/komposisi", result.get("Komposisi/Dosis", "Belum terdeteksi")],
        [result.get("Nama Tanaman", "Tanaman"), "bersumber dari", result.get("Sumber Data", "Belum terdeteksi")],
    ], columns=["Entitas Sumber", "Relasi", "Entitas Tujuan"])

    st.dataframe(rel_df)


def render_kg_section(result):
    html_block('<div class="section-title">🕸️ HerbKG 2.0 Explorer</div>')
    try:
        st.plotly_chart(make_kg_graph(result))
    except Exception as e:
        st.warning("Knowledge Graph belum dapat ditampilkan.")
        st.code(str(e))


def render_descriptive_chart():
    html_block('<div class="section-title">📊 Grafik Analisis Deskriptif</div>')

    fig = make_descriptive_chart(
        df_data,
        title="Analisis Deskriptif: 10 Tanaman dengan Data Terbanyak"
    )

    if fig is None:
        st.info("Grafik belum dapat dibuat karena kolom nama tanaman tidak ditemukan.")
    else:
        st.plotly_chart(fig)


def render_all_outputs(result, image_path, dataset_status, doc_status, match_status):
    html_block(f"""
    <div class="lilac-card">
        <h3>📌 Status Sistem</h3>
        <p><b>Status Dataset:</b> {safe_text(dataset_status)}</p>
        <p><b>Status Dokumen:</b> {safe_text(doc_status)}</p>
        <p><b>Status Koneksi Entitas:</b> {safe_text(match_status)}</p>
    </div>
    """)

    render_result_cards(result)
    render_image_section(result, image_path)
    render_relation_table(result)
    render_kg_section(result)
    render_descriptive_chart()


def render_preview_kg():
    sample = st.session_state.last_result

    if sample is None:
        sample = {
            "Nama Tanaman": "Jahe",
            "Nama Latin": "Zingiber officinale",
            "Nama Lokal/Daerah": "Jahe",
            "Bagian Tanaman": "Rimpang",
            "Zat Bioaktif": "Gingerol",
            "Khasiat/Efek Terapeutik": "Anti-inflamasi",
            "Cara Pengolahan": "Direbus",
            "Komposisi/Dosis": "Secukupnya",
            "Sumber Data": "Dataset Herbal",
            "Gambar": "Belum terdeteksi",
        }

    st.subheader("PREVIEW KNOWLEDGE GRAPH")
    try:
        st.plotly_chart(make_kg_graph(sample))
    except Exception:
        st.info("Preview KG belum tersedia.")


# =========================================================
# DOWNSTREAM AKTIF
# =========================================================
def ds_set_view(view_name):
    st.session_state.downstream_view = view_name


def ds_plot_count(column_name, title):
    if df_data.empty:
        st.warning("Dataset belum terbaca.")
        return

    if column_name is None:
        st.warning("Kolom yang dibutuhkan belum ditemukan pada dataset.")
        return

    fig = make_descriptive_chart(df_data, column=column_name, title=title)

    if fig:
        st.plotly_chart(fig)

    temp = (
        df_data[column_name]
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .value_counts()
        .head(30)
        .reset_index()
    )

    temp.columns = ["Kategori", "Jumlah Data"]
    st.dataframe(temp)


def ds_relation_table(mode):
    cols = get_columns(df_data)

    nama = cols["nama"]
    senyawa = cols["senyawa"]
    khasiat = cols["khasiat"]
    sumber = cols["sumber"]
    bagian = cols["bagian"]

    if df_data.empty:
        st.warning("Dataset belum terbaca.")
        return

    if mode == "tanaman_senyawa":
        if nama is None or senyawa is None:
            st.warning("Kolom Nama Tanaman atau Senyawa Bioaktif belum ditemukan.")
            return

        rel_df = df_data[[nama, senyawa]].copy()
        rel_df.columns = ["Entitas Sumber", "Entitas Tujuan"]
        rel_df["Relasi"] = "mengandung senyawa bioaktif"
        rel_df = rel_df[["Entitas Sumber", "Relasi", "Entitas Tujuan"]]

    elif mode == "senyawa_khasiat":
        if senyawa is None or khasiat is None:
            st.warning("Kolom Senyawa Bioaktif atau Khasiat belum ditemukan.")
            return

        rel_df = df_data[[senyawa, khasiat]].copy()
        rel_df.columns = ["Entitas Sumber", "Entitas Tujuan"]
        rel_df["Relasi"] = "mendukung khasiat"
        rel_df = rel_df[["Entitas Sumber", "Relasi", "Entitas Tujuan"]]

    elif mode == "khasiat_bukti":
        if khasiat is None:
            st.warning("Kolom Khasiat belum ditemukan.")
            return

        if sumber:
            rel_df = df_data[[khasiat, sumber]].copy()
            rel_df.columns = ["Khasiat", "Bukti / Sumber Data"]
        else:
            rel_df = df_data[[khasiat]].copy()
            rel_df.columns = ["Khasiat"]
            rel_df["Bukti / Sumber Data"] = "Dataset herbal"

        rel_df["Relasi"] = "didukung oleh bukti"
        rel_df = rel_df[["Khasiat", "Relasi", "Bukti / Sumber Data"]]

    else:
        needed = [c for c in [nama, bagian, senyawa, khasiat, sumber] if c is not None]

        if not needed:
            st.warning("Kolom relasi belum ditemukan pada dataset.")
            return

        rel_df = df_data[needed].copy()

        rename_map = {}
        if nama:
            rename_map[nama] = "Tanaman"
        if bagian:
            rename_map[bagian] = "Bagian Digunakan"
        if senyawa:
            rename_map[senyawa] = "Senyawa Bioaktif"
        if khasiat:
            rename_map[khasiat] = "Khasiat"
        if sumber:
            rename_map[sumber] = "Sumber Data"

        rel_df = rel_df.rename(columns=rename_map)

    rel_df = rel_df.fillna("")
    rel_df = rel_df.drop_duplicates().head(50)
    st.dataframe(rel_df)


def ds_similarity_analysis(target_name):
    cols = get_columns(df_data)

    nama = cols["nama"]
    senyawa = cols["senyawa"]
    khasiat = cols["khasiat"]
    bagian = cols["bagian"]

    if df_data.empty:
        st.warning("Dataset belum terbaca.")
        return

    if nama is None:
        st.warning("Kolom Nama Tanaman belum ditemukan.")
        return

    feature_cols = [c for c in [senyawa, khasiat, bagian] if c is not None]

    if not feature_cols:
        st.warning("Kolom fitur kemiripan belum ditemukan.")
        return

    target_clean = clean_text(target_name)
    target_row = None

    for _, row in df_data.iterrows():
        nama_row = clean_text(row.get(nama, ""))
        if target_clean in nama_row or nama_row in target_clean:
            target_row = row
            break

    if target_row is None:
        st.warning(f"Tanaman {target_name} belum ditemukan pada dataset.")
        return

    base_text = " ".join([str(target_row.get(c, "")) for c in feature_cols])
    base_tokens = set([t for t in clean_text(base_text).split() if len(t) >= 3])

    scores = []

    for _, row in df_data.iterrows():
        nama_row = str(row.get(nama, "")).strip()

        if clean_text(nama_row) == clean_text(target_name):
            continue

        row_text = " ".join([str(row.get(c, "")) for c in feature_cols])
        row_tokens = set([t for t in clean_text(row_text).split() if len(t) >= 3])

        if not base_tokens or not row_tokens:
            score = 0
        else:
            score = len(base_tokens.intersection(row_tokens)) / len(base_tokens.union(row_tokens))

        if score > 0:
            scores.append({
                "Tanaman Referensi": target_name,
                "Tanaman Mirip": nama_row,
                "Skor Kemiripan": round(score, 3)
            })

    result_df = pd.DataFrame(scores)

    if result_df.empty:
        st.info("Belum ditemukan tanaman yang mirip berdasarkan fitur dataset.")
        return

    result_df = result_df.sort_values("Skor Kemiripan", ascending=False).head(10)

    st.dataframe(result_df)

    fig = px.bar(
        result_df,
        x="Skor Kemiripan",
        y="Tanaman Mirip",
        orientation="h",
        text="Skor Kemiripan",
        color="Skor Kemiripan",
        color_continuous_scale=["#bbf7d0", "#22c55e", "#047857"],
        title=f"Analisis Kemiripan Tanaman terhadap {target_name}"
    )

    fig.update_layout(
        height=430,
        plot_bgcolor="white",
        paper_bgcolor="white",
        yaxis=dict(autorange="reversed"),
        margin=dict(l=20, r=20, t=60, b=20)
    )

    st.plotly_chart(fig)


def ds_recommendation(mode):
    cols = get_columns(df_data)

    nama = cols["nama"]
    senyawa = cols["senyawa"]
    khasiat = cols["khasiat"]
    sumber = cols["sumber"]

    if df_data.empty:
        st.warning("Dataset belum terbaca.")
        return

    if nama is None:
        st.warning("Kolom Nama Tanaman belum ditemukan.")
        return

    if mode == "peringkat":
        if khasiat:
            rank_df = (
                df_data.groupby(nama)[khasiat]
                .count()
                .reset_index()
                .rename(columns={nama: "Tanaman", khasiat: "Skor Rekomendasi"})
                .sort_values("Skor Rekomendasi", ascending=False)
                .head(10)
            )
        else:
            rank_df = df_data[nama].value_counts().head(10).reset_index()
            rank_df.columns = ["Tanaman", "Skor Rekomendasi"]

        st.dataframe(rank_df)

        fig = px.bar(
            rank_df,
            x="Skor Rekomendasi",
            y="Tanaman",
            orientation="h",
            text="Skor Rekomendasi",
            color="Skor Rekomendasi",
            color_continuous_scale=["#bbf7d0", "#22c55e", "#047857"],
            title="Peringkat Rekomendasi Tanaman Herbal"
        )

        fig.update_layout(
            height=430,
            plot_bgcolor="white",
            paper_bgcolor="white",
            yaxis=dict(autorange="reversed")
        )

        st.plotly_chart(fig)
        return

    if mode == "graph_search":
        st.info("Graph Search menelusuri hubungan Tanaman → Senyawa → Khasiat → Sumber Data.")
        ds_relation_table("jalur_relasi")
        return

    keyword = st.text_input(
        "Masukkan keluhan / penyakit / khasiat yang dicari",
        placeholder="Contoh: batuk, demam, diabetes, antiinflamasi"
    )

    search_cols = [c for c in [nama, senyawa, khasiat, sumber] if c is not None]

    if not keyword:
        st.info("Masukkan kata kunci untuk mencari rekomendasi herbal.")
        return

    temp = df_data.copy()
    mask = pd.Series(False, index=temp.index)

    for c in search_cols:
        mask = mask | temp[c].astype(str).str.lower().str.contains(keyword.lower(), na=False)

    hasil = temp[mask]
    tampil_cols = [c for c in [nama, senyawa, khasiat, sumber] if c is not None]

    if hasil.empty:
        st.info("Belum ditemukan rekomendasi berdasarkan kata kunci tersebut.")
    else:
        st.success(f"Ditemukan {len(hasil)} kandidat tanaman herbal.")
        st.dataframe(hasil[tampil_cols].drop_duplicates().head(50))


def render_downstream_result():
    view = st.session_state.downstream_view
    st.markdown("---")

    if view == "awal":
        st.info("Klik salah satu tombol downstream untuk menampilkan hasil analisis.")
        return

    cols = get_columns(df_data)

    if view == "deskriptif_tanaman":
        st.subheader("📊 Analisis Deskriptif: Tanaman")
        ds_plot_count(cols["nama"], "Distribusi Data Berdasarkan Nama Tanaman")

    elif view == "deskriptif_senyawa":
        st.subheader("🧪 Analisis Deskriptif: Senyawa Bioaktif")
        ds_plot_count(cols["senyawa"], "Distribusi Data Berdasarkan Senyawa Bioaktif")

    elif view == "deskriptif_khasiat":
        st.subheader("💚 Analisis Deskriptif: Khasiat")
        ds_plot_count(cols["khasiat"], "Distribusi Data Berdasarkan Khasiat")

    elif view == "deskriptif_sumber":
        st.subheader("📚 Analisis Deskriptif: Sumber Data")
        ds_plot_count(cols["sumber"], "Distribusi Data Berdasarkan Sumber Data")

    elif view == "query_tanaman_senyawa":
        st.subheader("🔎 Query Graf: Tanaman → Senyawa")
        ds_relation_table("tanaman_senyawa")

    elif view == "query_senyawa_khasiat":
        st.subheader("🔎 Query Graf: Senyawa → Khasiat")
        ds_relation_table("senyawa_khasiat")

    elif view == "query_khasiat_bukti":
        st.subheader("🔎 Query Graf: Khasiat → Bukti Literatur")
        ds_relation_table("khasiat_bukti")

    elif view == "query_jalur_relasi":
        st.subheader("🔎 Output: Jalur Relasi")
        ds_relation_table("jalur_relasi")

    elif view == "similarity_kelor":
        st.subheader("🧬 Analisis Kemiripan: Kelor")
        ds_similarity_analysis("Kelor")

    elif view == "similarity_sirih":
        st.subheader("🧬 Analisis Kemiripan: Sirih")
        ds_similarity_analysis("Sirih")

    elif view == "similarity_jahe":
        st.subheader("🧬 Analisis Kemiripan: Jahe")
        ds_similarity_analysis("Jahe")

    elif view == "similarity_kayu_manis":
        st.subheader("🧬 Analisis Kemiripan: Kayu Manis")
        ds_similarity_analysis("Kayu Manis")

    elif view == "rekomendasi_keluhan":
        st.subheader("💊 Rekomendasi Herbal Berdasarkan Keluhan / Penyakit")
        ds_recommendation("keluhan")

    elif view == "rekomendasi_graph":
        st.subheader("💊 Rekomendasi Herbal: Graph Search")
        ds_recommendation("graph_search")

    elif view == "rekomendasi_tanaman":
        st.subheader("💊 Tanaman Herbal Terkait")
        ds_recommendation("tanaman_terkait")

    elif view == "rekomendasi_peringkat":
        st.subheader("💊 Peringkat Rekomendasi Herbal")
        ds_recommendation("peringkat")


def render_downstream_page():
    html_block('<div class="section-title">📦 Aplikasi Downstream</div>')

    st.write(
        "Aplikasi downstream memanfaatkan hasil ekstraksi entitas dan relasi dari HerbKG 2.0 "
        "untuk analisis deskriptif, query graf berbasis bukti, analisis kemiripan, dan rekomendasi herbal."
    )

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        html_block("""
        <div class="down-card">
            <h3>1. Analisis Deskriptif</h3>
            <p>Menampilkan ringkasan statistik entitas, relasi, dan distribusi data herbal.</p>
        </div>
        """)
        st.markdown('<div class="down-btn">', unsafe_allow_html=True)
        st.button("Tanaman", key="btn_ds_tanaman", on_click=ds_set_view, args=("deskriptif_tanaman",))
        st.button("Senyawa Bioaktif", key="btn_ds_senyawa", on_click=ds_set_view, args=("deskriptif_senyawa",))
        st.button("Khasiat", key="btn_ds_khasiat", on_click=ds_set_view, args=("deskriptif_khasiat",))
        st.button("Sumber Data", key="btn_ds_sumber", on_click=ds_set_view, args=("deskriptif_sumber",))
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        html_block("""
        <div class="down-card">
            <h3>2. Query Graf Berbasis Bukti</h3>
            <p>Menelusuri hubungan tanaman, senyawa, khasiat, dan sumber literatur.</p>
        </div>
        """)
        st.markdown('<div class="down-btn">', unsafe_allow_html=True)
        st.button("Tanaman → Senyawa", key="btn_query_ts", on_click=ds_set_view, args=("query_tanaman_senyawa",))
        st.button("Senyawa → Khasiat", key="btn_query_sk", on_click=ds_set_view, args=("query_senyawa_khasiat",))
        st.button("Khasiat → Bukti Literatur", key="btn_query_kb", on_click=ds_set_view, args=("query_khasiat_bukti",))
        st.button("Output: Jalur Relasi", key="btn_query_jalur", on_click=ds_set_view, args=("query_jalur_relasi",))
        st.markdown('</div>', unsafe_allow_html=True)

    with col3:
        html_block("""
        <div class="down-card">
            <h3>3. Analisis Kemiripan</h3>
            <p>Menemukan tanaman yang mirip berdasarkan senyawa dan khasiat.</p>
        </div>
        """)
        st.markdown('<div class="down-btn">', unsafe_allow_html=True)
        st.button("Kelor → 0,56", key="btn_sim_kelor", on_click=ds_set_view, args=("similarity_kelor",))
        st.button("Sirih → 0,41", key="btn_sim_sirih", on_click=ds_set_view, args=("similarity_sirih",))
        st.button("Jahe → 0,39", key="btn_sim_jahe", on_click=ds_set_view, args=("similarity_jahe",))
        st.button("Kayu Manis → 0,28", key="btn_sim_kayu", on_click=ds_set_view, args=("similarity_kayu_manis",))
        st.markdown('</div>', unsafe_allow_html=True)

    with col4:
        html_block("""
        <div class="down-card">
            <h3>4. Rekomendasi Herbal</h3>
            <p>Memberikan rekomendasi tanaman herbal berbasis khasiat dan relasi graf.</p>
        </div>
        """)
        st.markdown('<div class="down-btn">', unsafe_allow_html=True)
        st.button("Keluhan / Penyakit", key="btn_rec_keluhan", on_click=ds_set_view, args=("rekomendasi_keluhan",))
        st.button("Graph Search", key="btn_rec_graph", on_click=ds_set_view, args=("rekomendasi_graph",))
        st.button("Tanaman Terkait", key="btn_rec_tanaman", on_click=ds_set_view, args=("rekomendasi_tanaman",))
        st.button("Peringkat Rekomendasi", key="btn_rec_rank", on_click=ds_set_view, args=("rekomendasi_peringkat",))
        st.markdown('</div>', unsafe_allow_html=True)

    render_downstream_result()


def render_downstream_preview():
    html_block("""
    <div class="section-title">📦 Aplikasi Downstream</div>
    <div class="result-card">
        <h4>1. Analisis Deskriptif</h4>
        <p>Ringkasan statistik tanaman, senyawa, khasiat, dan sumber data.</p>
    </div>
    <div class="result-card">
        <h4>2. Query Graf Berbasis Bukti</h4>
        <p>Menelusuri hubungan tanaman, senyawa, khasiat, dan literatur.</p>
    </div>
    <div class="result-card">
        <h4>3. Analisis Kemiripan</h4>
        <p>Membandingkan tanaman berdasarkan senyawa dan khasiat.</p>
    </div>
    <div class="result-card">
        <h4>4. Rekomendasi Herbal</h4>
        <p>Rekomendasi tanaman herbal berdasarkan khasiat dan relasi graf.</p>
    </div>
    """)


# =========================================================
# HALAMAN UTAMA
# =========================================================
if menu == "🏠 Dashboard":
    top_left, top_right = st.columns([3.2, 1])

    with top_left:
        html_block("""
        <div class="hero-banner">
            <h2>Selamat datang di HyTBIONEX</h2>
            <p>
            Platform cerdas untuk isolasi informasi bioaktif dari tanaman herbal Indonesia
            menggunakan pendekatan Hybrid Transformer serta integrasi HerbKG 2.0.
            </p>
        </div>
        """)

    with top_right:
        render_summary_card()

    render_input_area()

    down_left, down_right = st.columns([2, 1])

    with down_left:
        render_downstream_preview()

    with down_right:
        render_preview_kg()

    f1, f2, f3 = st.columns(3)

    with f1:
        st.success("⚙️ Model Aktif: Hybrid Transformer")

    with f2:
        st.success("🧬 Pipeline: NED → BIE → RE → HerbKG 2.0")

    with f3:
        st.success("🛡️ Status: Sistem siap digunakan")


elif menu in ["🌿 Input Tanaman", "📁 Upload Dokumen"]:
    render_input_area()


elif menu == "📋 Hasil Isolasi Entitas":
    if st.session_state.last_result:
        render_result_cards(st.session_state.last_result)
        render_image_section(st.session_state.last_result, st.session_state.last_image)
    else:
        st.warning("Belum ada hasil ekstraksi. Silakan proses data terlebih dahulu di menu Dashboard atau Input Tanaman.")


elif menu == "🔗 Relation Extraction":
    if st.session_state.last_result:
        render_relation_table(st.session_state.last_result)
    else:
        st.warning("Belum ada hasil relasi. Silakan proses data terlebih dahulu.")


elif menu == "🕸️ HerbKG 2.0 Explorer":
    if st.session_state.last_result:
        render_kg_section(st.session_state.last_result)
    else:
        render_preview_kg()


elif menu == "📦 Aplikasi Downstream":
    render_downstream_page()


elif menu == "📊 Statistik & Analitik":
    render_descriptive_chart()
    html_block('<div class="section-title">📋 Cuplikan Dataset</div>')
    safe_dataframe(df_data, n=30)


elif menu == "⚙️ Pengaturan":
    html_block("""
    <div class="lilac-card">
        <h2>⚙️ Pengaturan Sistem</h2>
        <p>Halaman ini disiapkan untuk pengaturan model, dataset, pipeline, dan tampilan sistem.</p>
    </div>
    """)


elif menu == "ℹ️ Tentang Aplikasi":
    html_block("""
    <div class="lilac-card">
        <h2>ℹ️ Tentang HyTBIONEX</h2>
        <p>
        HyTBIONEX adalah prototipe sistem ekstraksi informasi bioaktif tanaman herbal Indonesia
        berbasis pipeline Hybrid Transformer, ekstraksi entitas, ekstraksi relasi,
        visualisasi HerbKG 2.0, serta analisis deskriptif.
        </p>
        <p><b>Researcher:</b> Nazwita</p>
    </div>
    """)
