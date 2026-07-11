import os
import re
import html
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px


# =========================================================
# KONFIGURASI
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
    if df.empty:
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


def path_exists(path):
    if path and os.path.exists(path):
        return path
    return None


# =========================================================
# LOAD DATASET UTAMA
# =========================================================
@st.cache_data(show_spinner=False)
def load_dataset():
    if not os.path.exists(DATASET_FILE):
        excel_files = [
            f for f in os.listdir(".")
            if f.lower().endswith((".xlsx", ".xls")) and f != IMAGE_TABLE_FILE
        ]

        if excel_files:
            dataset_path = excel_files[0]
        else:
            return pd.DataFrame(), "Dataset Excel utama belum ditemukan."
    else:
        dataset_path = DATASET_FILE

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


# =========================================================
# LOAD DATA GAMBAR TANAMAN
# =========================================================
@st.cache_data(show_spinner=False)
def load_image_mapping():
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

        if not image_df.empty:
            image_df.columns = [str(c).strip() for c in image_df.columns]

            col_nama = find_col(image_df, ["Nama Tanaman", "Tanaman", "Nama"])
            col_latin = find_col(image_df, ["Nama Latin", "Latin"])
            col_gambar = find_col(image_df, ["Gambar", "Image", "Foto", "File Gambar", "Path Gambar"])

            for _, row in image_df.iterrows():
                nama = value_from_row(row, col_nama)
                latin = value_from_row(row, col_latin)
                gambar = value_from_row(row, col_gambar)

                keys = []

                if nama != "Belum terdeteksi":
                    keys.append(clean_text(nama))

                if latin != "Belum terdeteksi":
                    keys.append(clean_text(latin))

                for key in keys:
                    if key:
                        mapping[key] = {
                            "nama": nama,
                            "latin": latin,
                            "gambar": "" if gambar == "Belum terdeteksi" else gambar,
                            "embedded": ""
                        }

        # Ekstrak gambar tertanam dari Excel tambahan
        try:
            from openpyxl import load_workbook

            wb = load_workbook(IMAGE_TABLE_FILE)
            temp_dir = Path(tempfile.gettempdir()) / "hytbionex_embedded_images"
            temp_dir.mkdir(parents=True, exist_ok=True)

            embedded_paths = []

            for ws in wb.worksheets:
                images = getattr(ws, "_images", [])

                for idx, img in enumerate(images):
                    img_data = img._data()
                    out_path = temp_dir / f"embedded_{ws.title}_{idx}.png"

                    with open(out_path, "wb") as f:
                        f.write(img_data)

                    embedded_paths.append(str(out_path))

            # Khusus kalau hanya ada 1 gambar tertanam dan ada baris Jahe,
            # gambar otomatis dikaitkan ke Jahe.
            if len(embedded_paths) == 1:
                for key in list(mapping.keys()):
                    if "jahe" in key or "zingiber" in key:
                        mapping[key]["embedded"] = embedded_paths[0]

        except Exception:
            pass

    except Exception:
        pass

    return mapping


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
# EKSTRAKSI / MATCHING DATASET
# =========================================================
def score_row(row, search_text, col_nama, col_latin, col_lokal):
    score = 0

    nama = clean_text(value_from_row(row, col_nama))
    latin = clean_text(value_from_row(row, col_latin))
    lokal = clean_text(value_from_row(row, col_lokal))

    if nama and nama != "belum terdeteksi":
        if nama in search_text:
            score += 170

        for token in nama.split():
            if len(token) >= 3 and token in search_text:
                score += 25

    if latin and latin != "belum terdeteksi":
        if latin in search_text:
            score += 150

        for token in latin.split():
            if len(token) >= 4 and token in search_text:
                score += 20

    if lokal and lokal != "belum terdeteksi":
        parts = re.split(r"[,;/|]", lokal)

        for p in parts:
            p = clean_text(p)
            if p and len(p) >= 3 and p in search_text:
                score += 80

    return score


def find_best_match(df, search_text):
    if df.empty:
        return None, "Dataset belum terbaca."

    search_text = clean_text(search_text)

    if not search_text:
        return None, "Input tanaman dan dokumen masih kosong."

    col_nama = find_col(df, ["Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"])
    col_latin = find_col(df, ["Nama Latin", "Nama_Latin", "Latin"])
    col_lokal = find_col(df, ["Nama Lokal/Daerah", "Nama Lokal", "Nama Daerah", "Bahasa Daerah", "Bahasa_Daerah"])

    best_row = None
    best_score = 0

    for _, row in df.iterrows():
        score = score_row(row, search_text, col_nama, col_latin, col_lokal)

        if score > best_score:
            best_score = score
            best_row = row

    if best_row is not None and best_score > 0:
        nama = value_from_row(best_row, col_nama)
        latin = value_from_row(best_row, col_latin)

        return best_row, f"Entitas cocok dengan dataset: {nama} / {latin} | Skor: {best_score}"

    return None, "Tidak ditemukan kecocokan entitas tanaman pada dataset."


def extract_result(row, input_text, df):
    col_nama = find_col(df, ["Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"])
    col_latin = find_col(df, ["Nama Latin", "Nama_Latin", "Latin"])
    col_lokal = find_col(df, ["Nama Lokal/Daerah", "Nama Lokal", "Nama Daerah", "Bahasa Daerah", "Bahasa_Daerah"])
    col_bagian = find_col(df, ["Bagian Tanaman", "Bagian Digunakan", "Bagian_Digunakan", "Bagian"])
    col_senyawa = find_col(df, ["Zat Bioaktif", "Senyawa Bioaktif", "Senyawa_Bioaktif", "Compound", "Senyawa", "Kandungan", "Kandungan Kimia"])
    col_khasiat = find_col(df, ["Khasiat/Efek Terapeutik", "Khasiat", "Manfaat", "Benefit", "Biological Activity", "Biological_Activity"])
    col_pengolahan = find_col(df, ["Cara Pengolahan", "Cara_Pengolahan", "Pengolahan", "Cara Pemakaian"])
    col_dosis = find_col(df, ["Komposisi/Dosis", "Komposisi /Dosis", "Dosis", "Komposisi"])
    col_sumber = find_col(df, ["Sumber Data", "Sumber_Data", "Sumber", "Referensi"])
    col_gambar = find_col(df, ["Gambar", "Image", "Foto", "File Gambar", "Path Gambar", "Nama File Gambar"])

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
        "Nama Tanaman": value_from_row(row, col_nama),
        "Nama Latin": value_from_row(row, col_latin),
        "Nama Lokal/Daerah": value_from_row(row, col_lokal),
        "Bagian Tanaman": value_from_row(row, col_bagian),
        "Zat Bioaktif": value_from_row(row, col_senyawa),
        "Khasiat/Efek Terapeutik": value_from_row(row, col_khasiat),
        "Cara Pengolahan": value_from_row(row, col_pengolahan),
        "Komposisi/Dosis": value_from_row(row, col_dosis),
        "Sumber Data": value_from_row(row, col_sumber),
        "Gambar": value_from_row(row, col_gambar),
    }


# =========================================================
# GAMBAR TANAMAN
# =========================================================
def find_plant_image(result, image_mapping):
    gambar = result.get("Gambar", "")
    nama = result.get("Nama Tanaman", "")
    latin = result.get("Nama Latin", "")

    candidates = []

    # Dari kolom Gambar dataset utama
    if gambar and gambar != "Belum terdeteksi":
        candidates.extend([
            gambar,
            os.path.join(ASSET_DIR, gambar),
            os.path.join("gambar", gambar),
            os.path.join("images", gambar),
        ])

    # Dari file Gambar tanaman herbal.xlsx
    map_keys = [clean_text(nama), clean_text(latin)]

    for key in map_keys:
        if key in image_mapping:
            map_item = image_mapping[key]

            map_gambar = map_item.get("gambar", "")
            map_embedded = map_item.get("embedded", "")

            if map_gambar:
                candidates.extend([
                    map_gambar,
                    os.path.join(ASSET_DIR, map_gambar),
                    os.path.join("gambar", map_gambar),
                    os.path.join("images", map_gambar),
                ])

            if map_embedded:
                candidates.append(map_embedded)

    # Dari nama file otomatis
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
        if path_exists(path):
            return path

    return None


def run_extraction(text_input, uploaded_file, df, dataset_status, image_mapping):
    doc_text, doc_status = read_uploaded_file(uploaded_file)
    combined_text = f"{text_input} {doc_text}"

    row, match_status = find_best_match(df, combined_text)
    result = extract_result(row, text_input, df)
    image_path = find_plant_image(result, image_mapping)

    st.session_state.last_result = result
    st.session_state.last_image = image_path
    st.session_state.last_status = {
        "dataset": dataset_status,
        "document": doc_status,
        "match": match_status,
    }

    return result, image_path, doc_status, match_status


# =========================================================
# GRAFIK DESKRIPTIF DAN KG
# =========================================================
def make_descriptive_chart(df):
    if df.empty:
        return None

    col_nama = find_col(df, ["Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"])

    if col_nama is None:
        return None

    chart_df = (
        df[col_nama]
        .astype(str)
        .str.strip()
        .replace("", pd.NA)
        .dropna()
        .value_counts()
        .head(10)
        .reset_index()
    )

    chart_df.columns = ["Nama Tanaman", "Jumlah Data"]

    fig = px.bar(
        chart_df,
        x="Jumlah Data",
        y="Nama Tanaman",
        orientation="h",
        text="Jumlah Data",
        color="Jumlah Data",
        color_continuous_scale=["#bbf7d0", "#22c55e", "#047857"],
        title="Analisis Deskriptif: 10 Tanaman dengan Data Terbanyak"
    )

    fig.update_layout(
        height=430,
        plot_bgcolor="white",
        paper_bgcolor="white",
        title_font=dict(size=20, color="#064e3b"),
        xaxis_title="Jumlah Data",
        yaxis_title="Nama Tanaman",
        margin=dict(l=20, r=20, t=60, b=20),
    )

    fig.update_traces(textposition="outside")
    fig.update_yaxes(autorange="reversed")

    return fig


def short_label(text, max_len=24):
    text = str(text)

    if text in ["", "nan", "None", "Belum terdeteksi"]:
        return "Belum terdeteksi"

    if len(text) <= max_len:
        return text

    return text[:max_len] + "..."


def make_kg_graph(result):
    tanaman = short_label(result["Nama Tanaman"], 18)
    latin = short_label(result["Nama Latin"], 24)
    bagian = short_label(result["Bagian Tanaman"], 18)
    senyawa = short_label(result["Zat Bioaktif"], 24)
    khasiat = short_label(result["Khasiat/Efek Terapeutik"], 22)
    pengolahan = short_label(result["Cara Pengolahan"], 22)
    dosis = short_label(result["Komposisi/Dosis"], 22)
    sumber = short_label(result["Sumber Data"], 26)

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
# CSS
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

/* DASHBOARD */
.hero-banner {
    background:
        radial-gradient(circle at 70% 40%, rgba(187,247,208,0.92), transparent 28%),
        linear-gradient(135deg, #ffffff 0%, #ecfdf5 55%, #f8fafc 100%);
    background-size: cover;
    background-position: center;
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

.downstream-card {
    background: #ffffff;
    border-radius: 20px;
    padding: 18px;
    box-shadow: 0 10px 24px rgba(15,23,42,0.08);
    border: 1px solid rgba(6,78,59,0.12);
    min-height: 360px;
}

.downstream-card h4 {
    color: #064e3b;
    font-weight: 900;
    text-align: center;
}

.downstream-card p {
    color: #334155;
    font-size: 14px;
    line-height: 1.5;
}

.flow-box {
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    color: #064e3b;
    border-radius: 12px;
    padding: 8px;
    text-align: center;
    margin-bottom: 8px;
    font-weight: 800;
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
# LOAD DATA
# =========================================================
df_data, dataset_status = load_dataset()
image_mapping = load_image_mapping()

col_senyawa = find_col(df_data, ["Zat Bioaktif", "Senyawa Bioaktif", "Senyawa", "Compound", "Kandungan"])
col_khasiat = find_col(df_data, ["Khasiat", "Manfaat", "Benefit", "Khasiat/Efek Terapeutik", "Biological Activity"])

total_data = len(df_data)
total_senyawa = df_data[col_senyawa].astype(str).replace("", pd.NA).dropna().nunique() if col_senyawa else 0
total_khasiat = df_data[col_khasiat].astype(str).replace("", pd.NA).dropna().nunique() if col_khasiat else 0
total_relasi = total_data * 8 if total_data else 0


# =========================================================
# SIDEBAR CUSTOM
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
        if st.sidebar.button(label, key=key, use_container_width=True):
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
# HEADER ATAS
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
    with st.container():
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
        <p>Masukkan nama tanaman atau kalimat, lalu unggah dokumen jika ada. Sistem akan mencocokkan informasi dengan dataset herbal.</p>
    </div>
    """)

    col1, col2 = st.columns([1.1, 1])

    with col1:
        text_input = st.text_area(
            "Input Data Tanaman",
            placeholder="Contoh: Jahe, Kunyit, Sambiloto, Kayu Manis...",
            height=130,
            key="main_input_text"
        )

    with col2:
        uploaded_file = st.file_uploader(
            "Upload Dokumen",
            type=["pdf", "txt", "csv", "xlsx", "xls"],
            key="main_uploaded_file"
        )

    submit = st.button("🔍 Proses Analisis", use_container_width=True, key="btn_process_all")

    if submit:
        if not text_input and uploaded_file is None:
            st.warning("Masukkan nama tanaman atau upload dokumen terlebih dahulu.")
        else:
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
        ("🌿 Nama Tanaman", result["Nama Tanaman"]),
        ("🔬 Nama Latin", result["Nama Latin"]),
        ("🇮🇩 Nama Lokal/Daerah", result["Nama Lokal/Daerah"]),
        ("🍃 Bagian Tanaman", result["Bagian Tanaman"]),
        ("🧪 Zat Bioaktif", result["Zat Bioaktif"]),
        ("💚 Khasiat / Efek Terapeutik", result["Khasiat/Efek Terapeutik"]),
        ("☕ Cara Pengolahan", result["Cara Pengolahan"]),
        ("⚖️ Komposisi / Dosis", result["Komposisi/Dosis"]),
        ("📚 Sumber Data", result["Sumber Data"]),
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

    if image_path:
        c1, c2 = st.columns([1, 2])

        with c1:
            st.image(image_path, caption=result["Nama Tanaman"], use_container_width=True)

        with c2:
            html_block(f"""
            <div class="lilac-card">
                <h3>🌿 Gambar Tanaman Terkoneksi Dataset</h3>
                <p><b>Nama Tanaman:</b> {safe_text(result["Nama Tanaman"])}</p>
                <p><b>Nama Latin:</b> {safe_text(result["Nama Latin"])}</p>
                <p><b>Catatan:</b> Gambar ditampilkan dari kolom <b>Gambar</b> pada dataset, folder <b>assets</b>, atau file <b>Gambar tanaman herbal.xlsx</b>.</p>
            </div>
            """)
    else:
        st.info("Gambar belum ditemukan. Isi kolom Gambar di Excel, contoh: assets/jahe.jpg atau jahe.jpg.")


def render_relation_table(result):
    html_block('<div class="section-title">🔗 Relation Extraction</div>')

    rel_df = pd.DataFrame([
        [result["Nama Tanaman"], "memiliki nama latin", result["Nama Latin"]],
        [result["Nama Tanaman"], "memiliki nama lokal/daerah", result["Nama Lokal/Daerah"]],
        [result["Nama Tanaman"], "menggunakan bagian tanaman", result["Bagian Tanaman"]],
        [result["Nama Tanaman"], "mengandung senyawa bioaktif", result["Zat Bioaktif"]],
        [result["Nama Tanaman"], "memiliki khasiat", result["Khasiat/Efek Terapeutik"]],
        [result["Nama Tanaman"], "diolah dengan cara", result["Cara Pengolahan"]],
        [result["Nama Tanaman"], "memiliki dosis/komposisi", result["Komposisi/Dosis"]],
        [result["Nama Tanaman"], "bersumber dari", result["Sumber Data"]],
    ], columns=["Entitas Sumber", "Relasi", "Entitas Tujuan"])

    st.dataframe(rel_df, use_container_width=True)


def render_kg_section(result):
    html_block('<div class="section-title">🕸️ HerbKG 2.0 Explorer</div>')
    st.plotly_chart(make_kg_graph(result), use_container_width=True)


def render_descriptive_chart():
    html_block('<div class="section-title">📊 Grafik Analisis Deskriptif</div>')

    fig = make_descriptive_chart(df_data)

    if fig is None:
        st.info("Grafik belum dapat dibuat karena kolom nama tanaman tidak ditemukan.")
    else:
        st.plotly_chart(fig, use_container_width=True)


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
    st.plotly_chart(make_kg_graph(sample), use_container_width=True)


def render_downstream_page():
    html_block('<div class="section-title">📦 Aplikasi Downstream</div>')

    st.write(
        "Aplikasi downstream memanfaatkan hasil ekstraksi entitas dan relasi dari HerbKG 2.0 "
        "untuk analisis deskriptif, query graf berbasis bukti, analisis kemiripan, dan rekomendasi herbal."
    )

    d1, d2, d3, d4 = st.columns(4)

    with d1:
        html_block("""
        <div class="downstream-card">
            <h4>1. Analisis Deskriptif</h4>
            <p>Menampilkan ringkasan statistik entitas, relasi, dan distribusi data herbal.</p>
            <div class="flow-box">Tanaman</div>
            <div class="flow-box">Senyawa Bioaktif</div>
            <div class="flow-box">Khasiat</div>
            <div class="flow-box">Sumber Data</div>
        </div>
        """)

    with d2:
        html_block("""
        <div class="downstream-card">
            <h4>2. Query Graf Berbasis Bukti</h4>
            <p>Menelusuri hubungan tanaman, senyawa, khasiat, dan sumber literatur.</p>
            <div class="flow-box">Tanaman → Senyawa</div>
            <div class="flow-box">Senyawa → Khasiat</div>
            <div class="flow-box">Khasiat → Bukti Literatur</div>
            <div class="flow-box">Output: Jalur Relasi</div>
        </div>
        """)

    with d3:
        html_block("""
        <div class="downstream-card">
            <h4>3. Analisis Kemiripan</h4>
            <p>Menemukan tanaman yang mirip berdasarkan senyawa dan khasiat.</p>
            <div class="flow-box">Kelor → 0,56</div>
            <div class="flow-box">Sirih → 0,41</div>
            <div class="flow-box">Jahe → 0,39</div>
            <div class="flow-box">Kayu Manis → 0,28</div>
        </div>
        """)

    with d4:
        html_block("""
        <div class="downstream-card">
            <h4>4. Rekomendasi Herbal</h4>
            <p>Memberikan rekomendasi tanaman herbal berbasis khasiat dan relasi graf.</p>
            <div class="flow-box">Keluhan / Penyakit</div>
            <div class="flow-box">Graph Search</div>
            <div class="flow-box">Tanaman Terkait</div>
            <div class="flow-box">Peringkat Rekomendasi</div>
        </div>
        """)

    render_descriptive_chart()


# =========================================================
# HALAMAN
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
        render_downstream_page()

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
    st.dataframe(df_data.head(30), use_container_width=True)


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
