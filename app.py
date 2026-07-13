import os
import re
import html
import base64
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

if "input_text" not in st.session_state:
    st.session_state.input_text = ""


# =========================================================
# FUNGSI DASAR
# =========================================================
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


def image_to_base64(path):
    try:
        with open(path, "rb") as file:
            encoded = base64.b64encode(file.read()).decode()
        ext = Path(path).suffix.lower().replace(".", "")
        if ext == "jpg":
            ext = "jpeg"
        return f"data:image/{ext};base64,{encoded}"
    except Exception:
        return ""


def find_optional_background():
    candidates = [
        "assets/background.png",
        "assets/background.jpg",
        "assets/herbal_banner.png",
        "assets/herbal_banner.jpg",
        "assets/dashboard_background.png",
        "assets/dashboard_background.jpg",
        "background.png",
        "background.jpg",
    ]

    for path in candidates:
        if os.path.exists(path):
            return image_to_base64(path)

    return ""


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


@st.cache_data(show_spinner=False)
def load_dataset():
    if not os.path.exists(DATASET_FILE):
        excel_files = [f for f in os.listdir(".") if f.lower().endswith((".xlsx", ".xls"))]
        if excel_files:
            dataset_path = excel_files[0]
        else:
            return pd.DataFrame(), "Dataset Excel belum ditemukan."
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


def find_plant_image(result):
    gambar = result.get("Gambar", "")
    nama = result.get("Nama Tanaman", "")

    candidates = []

    if gambar and gambar != "Belum terdeteksi":
        candidates.extend([
            gambar,
            os.path.join(ASSET_DIR, gambar),
            os.path.join("gambar", gambar),
            os.path.join("images", gambar),
        ])

    slug = slugify_filename(nama)

    if slug:
        for ext in ["jpg", "jpeg", "png", "webp"]:
            candidates.extend([
                os.path.join(ASSET_DIR, f"{slug}.{ext}"),
                os.path.join("gambar", f"{slug}.{ext}"),
                os.path.join("images", f"{slug}.{ext}"),
                f"{slug}.{ext}",
            ])

    for path in candidates:
        if path and os.path.exists(path):
            return path

    return None


def run_extraction(text_input, uploaded_file, df, dataset_status):
    doc_text, doc_status = read_uploaded_file(uploaded_file)
    combined_text = f"{text_input} {doc_text}"

    row, match_status = find_best_match(df, combined_text)
    result = extract_result(row, text_input, df)
    image_path = find_plant_image(result)

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
    """
    Membuat Knowledge Graph tanaman herbal dengan entitas:
    1. Nama Tanaman
    2. Nama Latin
    3. Nama Lokal/Daerah
    4. Bagian Tanaman
    5. Zat Bioaktif
    6. Khasiat/Efek Terapeutik
    7. Cara Pengolahan
    8. Komposisi/Dosis
    9. Sumber Data
    """

    # =====================================================
    # MENGAMBIL DATA HASIL EKSTRAKSI
    # =====================================================
    tanaman = short_label(
        result.get("Nama Tanaman", "Belum terdeteksi"),
        18
    )

    latin = short_label(
        result.get("Nama Latin", "Belum terdeteksi"),
        24
    )

    lokal = short_label(
        result.get("Nama Lokal/Daerah", "Belum terdeteksi"),
        22
    )

    bagian = short_label(
        result.get("Bagian Tanaman", "Belum terdeteksi"),
        20
    )

    senyawa = short_label(
        result.get("Zat Bioaktif", "Belum terdeteksi"),
        28
    )

    khasiat = short_label(
        result.get("Khasiat/Efek Terapeutik", "Belum terdeteksi"),
        26
    )

    pengolahan = short_label(
        result.get("Cara Pengolahan", "Belum terdeteksi"),
        26
    )

    dosis = short_label(
        result.get("Komposisi/Dosis", "Belum terdeteksi"),
        24
    )

    sumber = short_label(
        result.get("Sumber Data", "Belum terdeteksi"),
        27
    )

    # =====================================================
    # POSISI NODE KNOWLEDGE GRAPH
    # =====================================================
    nodes = [
        {
            "id": "tanaman",
            "label": tanaman,
            "full_text": result.get("Nama Tanaman", ""),
            "x": 0,
            "y": 0,
            "color": "#047857",
            "border": "#065f46",
            "size": 88
        },

        {
            "id": "latin",
            "label": latin,
            "full_text": result.get("Nama Latin", ""),
            "x": -2.8,
            "y": 1.55,
            "color": "#e9d5ff",
            "border": "#9333ea",
            "size": 66
        },

        {
            "id": "lokal",
            "label": lokal,
            "full_text": result.get("Nama Lokal/Daerah", ""),
            "x": 0,
            "y": 2.15,
            "color": "#bfdbfe",
            "border": "#2563eb",
            "size": 66
        },

        {
            "id": "bagian",
            "label": bagian,
            "full_text": result.get("Bagian Tanaman", ""),
            "x": 2.8,
            "y": 1.55,
            "color": "#bbf7d0",
            "border": "#047857",
            "size": 66
        },

        {
            "id": "senyawa",
            "label": senyawa,
            "full_text": result.get("Zat Bioaktif", ""),
            "x": -3.15,
            "y": 0,
            "color": "#fed7aa",
            "border": "#f97316",
            "size": 72
        },

        {
            "id": "pengolahan",
            "label": pengolahan,
            "full_text": result.get("Cara Pengolahan", ""),
            "x": 3.15,
            "y": 0,
            "color": "#bbf7d0",
            "border": "#047857",
            "size": 70
        },

        {
            "id": "dosis",
            "label": dosis,
            "full_text": result.get("Komposisi/Dosis", ""),
            "x": -2.45,
            "y": -1.75,
            "color": "#fde68a",
            "border": "#d97706",
            "size": 68
        },

        {
            "id": "khasiat",
            "label": khasiat,
            "full_text": result.get(
                "Khasiat/Efek Terapeutik",
                ""
            ),
            "x": 0,
            "y": -2.2,
            "color": "#fbcfe8",
            "border": "#ec4899",
            "size": 74
        },

        {
            "id": "sumber",
            "label": sumber,
            "full_text": result.get("Sumber Data", ""),
            "x": 2.45,
            "y": -1.75,
            "color": "#e9d5ff",
            "border": "#9333ea",
            "size": 68
        },
    ]

    node_map = {
        node["id"]: node
        for node in nodes
    }

    # =====================================================
    # RELASI ANTARENTITAS
    # =====================================================
    edges = [
        (
            "tanaman",
            "latin",
            "nama latin"
        ),
        (
            "tanaman",
            "lokal",
            "nama lokal/daerah"
        ),
        (
            "tanaman",
            "bagian",
            "bagian digunakan"
        ),
        (
            "tanaman",
            "senyawa",
            "mengandung"
        ),
        (
            "tanaman",
            "pengolahan",
            "cara pengolahan"
        ),
        (
            "tanaman",
            "dosis",
            "dosis/komposisi"
        ),
        (
            "tanaman",
            "khasiat",
            "memiliki khasiat"
        ),
        (
            "tanaman",
            "sumber",
            "sumber data"
        ),
    ]

    # Pergeseran tulisan relasi agar tidak bertumpuk
    label_shift = {
        "nama latin": (0, 12),
        "nama lokal/daerah": (45, 0),
        "bagian digunakan": (0, 12),
        "mengandung": (0, 14),
        "cara pengolahan": (0, 14),
        "dosis/komposisi": (-5, -2),
        "memiliki khasiat": (50, 0),
        "sumber data": (5, -2),
    }

    fig = go.Figure()

    # =====================================================
    # MEMBUAT GARIS DAN LABEL RELASI
    # =====================================================
    for source, target, relation_label in edges:
        source_node = node_map[source]
        target_node = node_map[target]

        # Garis penghubung
        fig.add_trace(
            go.Scatter(
                x=[
                    source_node["x"],
                    target_node["x"]
                ],
                y=[
                    source_node["y"],
                    target_node["y"]
                ],
                mode="lines",
                line=dict(
                    color="#94a3b8",
                    width=2.5
                ),
                hoverinfo="none",
                showlegend=False
            )
        )

        x_shift, y_shift = label_shift.get(
            relation_label,
            (0, 0)
        )

        # Label relasi
        fig.add_annotation(
            x=(
                source_node["x"]
                + target_node["x"]
            ) / 2,
            y=(
                source_node["y"]
                + target_node["y"]
            ) / 2,
            text=relation_label,
            showarrow=False,
            xshift=x_shift,
            yshift=y_shift,
            font=dict(
                size=12,
                color="#111827",
                family="Arial"
            ),
            bgcolor="rgba(255,255,255,0.92)",
            bordercolor="#dbe3ec",
            borderwidth=1,
            borderpad=4
        )

    # =====================================================
    # MEMBUAT NODE
    # =====================================================
    for node in nodes:
        hover_text = (
            f"<b>{node['label']}</b>"
            f"<br>{node['full_text']}"
        )

        fig.add_trace(
            go.Scatter(
                x=[node["x"]],
                y=[node["y"]],
                mode="markers+text",
                marker=dict(
                    size=node["size"],
                    color=node["color"],
                    line=dict(
                        color=node["border"],
                        width=4
                    )
                ),
                text=[node["label"]],
                textposition="middle center",
                textfont=dict(
                    size=13,
                    color="#111111",
                    family="Arial Black"
                ),
                hovertext=[hover_text],
                hoverinfo="text",
                showlegend=False,
                cliponaxis=False
            )
        )

    # =====================================================
    # TAMPILAN KNOWLEDGE GRAPH
    # =====================================================
    fig.update_layout(
        height=720,
        plot_bgcolor="#fbf7ff",
        paper_bgcolor="#fbf7ff",
        margin=dict(
            l=20,
            r=20,
            t=30,
            b=30
        ),
        xaxis=dict(
            visible=False,
            range=[-4.1, 4.1],
            fixedrange=True
        ),
        yaxis=dict(
            visible=False,
            range=[-2.9, 2.8],
            fixedrange=True
        ),
        dragmode=False,
        hoverlabel=dict(
            bgcolor="white",
            font_size=13,
            font_family="Arial"
        )
    )

    return fig
    return fig
    def render_kg_section(result):
    st.markdown(
        '<div class="section-title">'
        '🕸️ HerbKG 2.0 Explorer'
        '</div>',
        unsafe_allow_html=True
    )

    fig_kg = make_kg_graph(result)

    st.plotly_chart(
        fig_kg,
        use_container_width=True,
        config={
            "displayModeBar": False,
            "scrollZoom": False
        }
    )
# =========================================================
# CSS
# =========================================================
hero_bg = find_optional_background()

if hero_bg:
    hero_background_css = f"""
        linear-gradient(90deg, rgba(255,255,255,0.96), rgba(255,255,255,0.58)),
        url("{hero_bg}")
    """
else:
    hero_background_css = """
        radial-gradient(circle at 70% 40%, rgba(187,247,208,0.92), transparent 28%),
        linear-gradient(135deg, #ffffff 0%, #ecfdf5 55%, #f8fafc 100%)
    """

st.markdown(f"""
<style>
.stApp {{
    background: #f5fbf7 !important;
}}

.main .block-container {{
    padding-top: 1.2rem;
    max-width: 1500px;
}}

/* SIDEBAR */
[data-testid="stSidebar"] {{
    background:
        radial-gradient(circle at bottom left, rgba(34,197,94,0.25), transparent 28%),
        radial-gradient(circle at top right, rgba(16,185,129,0.18), transparent 35%),
        linear-gradient(180deg, #021f16 0%, #043b2c 45%, #065f46 100%) !important;
    border-right: 1px solid rgba(255,255,255,0.12);
}}

[data-testid="stSidebar"] * {{
    color: #f8fff8 !important;
}}

[data-testid="stSidebar"] h1 {{
    color: #ffffff !important;
    font-size: 34px !important;
    font-weight: 900 !important;
}}

[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {{
    color: #86efac !important;
    font-weight: 900 !important;
}}

[data-testid="stSidebar"] .stButton > button {{
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
}}

[data-testid="stSidebar"] .stButton > button:hover {{
    background: rgba(255,255,255,0.12) !important;
    color: #ffffff !important;
    transform: translateX(3px);
}}

.nav-active {{
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
}}

.sidebar-section-title {{
    color: #86efac !important;
    font-size: 13px !important;
    font-weight: 900 !important;
    letter-spacing: 0.6px;
    margin-top: 22px;
    margin-bottom: 10px;
    text-transform: uppercase;
}}

.sidebar-line {{
    height: 1px;
    background: rgba(255,255,255,0.14);
    margin: 18px 0;
}}

.sidebar-footer {{
    margin-top: 28px;
    padding: 18px;
    border-radius: 18px;
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.14);
}}

.sidebar-footer h3 {{
    color: #ffffff !important;
    margin: 0 0 6px 0 !important;
    font-size: 17px !important;
    font-weight: 900 !important;
}}

.sidebar-footer p {{
    color: #d1fae5 !important;
    margin: 0 !important;
    font-size: 12px !important;
}}

/* HEADER ATAS */
.top-header {{
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
}}

.top-header h1 {{
    margin: 0;
    font-size: 38px;
    font-weight: 900;
    color: #ffffff;
}}

.top-header h2 {{
    margin: 0;
    font-size: 28px;
    font-weight: 900;
    color: #bbf7d0;
}}

.top-header p {{
    margin: 5px 0 0 0;
    color: #ecfdf5;
}}

.top-user {{
    display: flex;
    gap: 12px;
    justify-content: flex-end;
    align-items: center;
}}

.top-pill {{
    background: rgba(255,255,255,0.10);
    padding: 11px 14px;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.12);
    color: #ffffff;
    font-weight: 800;
    font-size: 14px;
}}

.user-card {{
    background: rgba(255,255,255,0.10);
    padding: 10px 14px;
    border-radius: 16px;
    border: 1px solid rgba(255,255,255,0.12);
    color: #ffffff;
    font-weight: 800;
    min-width: 145px;
}}

/* DASHBOARD */
.dashboard-shell {{
    background: rgba(255,255,255,0.92);
    border-radius: 30px;
    padding: 22px;
    box-shadow: 0 18px 44px rgba(15,23,42,0.13);
    border: 1px solid rgba(6,78,59,0.10);
}}

.hero-banner {{
    background: {hero_background_css};
    background-size: cover;
    background-position: center;
    padding: 34px 32px;
    border-radius: 22px;
    min-height: 215px;
    box-shadow: inset 0 0 0 1px rgba(6,78,59,0.08);
}}

.hero-banner h2 {{
    color: #047857;
    font-size: 30px;
    font-weight: 900;
    margin-bottom: 14px;
}}

.hero-banner p {{
    color: #12372a;
    font-size: 17px;
    line-height: 1.65;
    max-width: 680px;
}}

.single-input-card {{
    background: rgba(255,255,255,0.97);
    padding: 25px;
    border-radius: 22px;
    box-shadow: 0 12px 28px rgba(15,23,42,0.10);
    border: 1px solid rgba(6,78,59,0.10);
    margin-top: 18px;
    margin-bottom: 18px;
}}

.single-input-card h3 {{
    color: #064e3b;
    margin: 0 0 10px 0;
    font-weight: 900;
}}

.single-input-card p {{
    color: #334155;
}}

textarea {{
    background: #ffffff !important;
    color: #0f172a !important;
    border: 1px solid #d1d5db !important;
    border-radius: 12px !important;
}}

textarea::placeholder {{
    color: #94a3b8 !important;
}}

[data-testid="stFileUploader"] section {{
    background: #fbfffb !important;
    border: 2px dashed #86efac !important;
    border-radius: 16px !important;
    min-height: 125px;
}}

[data-testid="stFileUploader"] section * {{
    color: #0f172a !important;
}}

[data-testid="stFileUploader"] button {{
    background: #ffffff !important;
    color: #047857 !important;
    border: 1px solid #bbf7d0 !important;
    border-radius: 12px !important;
    font-weight: 800 !important;
}}

.stButton > button {{
    background: linear-gradient(90deg, #047857, #059669) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 13px !important;
    font-weight: 900 !important;
    padding: 0.72rem 1rem !important;
    box-shadow: 0 8px 18px rgba(4,120,87,0.18);
}}

.stButton > button:hover {{
    background: linear-gradient(90deg, #065f46, #047857) !important;
    color: white !important;
}}

.summary-card {{
    background: rgba(255,255,255,0.96);
    padding: 22px;
    border-radius: 20px;
    box-shadow: 0 12px 28px rgba(15,23,42,0.10);
    border: 1px solid rgba(6,78,59,0.10);
}}

.summary-item {{
    display: flex;
    gap: 14px;
    align-items: center;
    margin-bottom: 22px;
}}

.summary-icon {{
    width: 48px;
    height: 48px;
    border-radius: 16px;
    display:flex;
    align-items:center;
    justify-content:center;
    font-size: 25px;
    background:#ecfdf5;
}}

.summary-label {{
    color:#047857;
    font-weight:800;
}}

.summary-value {{
    color:#16a34a;
    font-size:24px;
    font-weight:900;
}}

.quick-card {{
    background: rgba(255,255,255,0.96);
    padding: 18px;
    border-radius: 18px;
    border: 1px solid rgba(6,78,59,0.10);
    box-shadow: 0 10px 24px rgba(15,23,42,0.08);
    text-align:center;
    min-height: 205px;
}}

.quick-card h4 {{
    color: #064e3b;
    font-weight: 900;
}}

.quick-card p {{
    color: #334155;
    font-size: 14px;
}}

.quick-btn {{
    display:inline-block;
    margin-top:10px;
    padding:8px 18px;
    border-radius:999px;
    border:1px solid #86efac;
    color:#047857;
    font-weight:800;
    background:#f8fff8;
}}

.lilac-card {{
    background: #f3e8ff;
    padding: 22px;
    border-radius: 18px;
    border: 2px solid #c084fc;
    color: #111111;
    margin-bottom: 18px;
}}

.result-card {{
    background: #ffffff;
    border-left: 7px solid #047857;
    border-radius: 15px;
    padding: 16px;
    min-height: 112px;
    box-shadow: 0 8px 18px rgba(15,23,42,0.09);
    margin-bottom: 12px;
}}

.result-card h4 {{
    color: #064e3b;
    font-weight: 900;
    margin-bottom: 7px;
}}

.result-card p {{
    color: #0f172a;
    font-size: 16px;
    line-height: 1.5;
}}

.section-title {{
    color: #064e3b;
    font-size: 27px;
    font-weight: 900;
    margin-top: 22px;
    margin-bottom: 13px;
}}

.graph-box {{
    background: #ffffff;
    border-radius: 20px;
    padding: 18px;
    box-shadow: 0 10px 25px rgba(15,23,42,0.09);
    border: 1px solid rgba(6,78,59,0.10);
    margin-bottom: 18px;
}}

.footer-status {{
    background: rgba(255,255,255,0.94);
    border-radius: 16px;
    padding: 17px;
    border: 1px solid rgba(6,78,59,0.12);
    color:#064e3b;
    font-weight:800;
    text-align:center;
}}

.downstream-card {{
    background: #ffffff;
    border-radius: 20px;
    padding: 18px;
    box-shadow: 0 10px 24px rgba(15,23,42,0.08);
    border: 1px solid rgba(6,78,59,0.12);
    min-height: 360px;
}}

.downstream-card h4 {{
    color: #064e3b;
    font-weight: 900;
    text-align: center;
}}

.downstream-card p {{
    color: #334155;
    font-size: 14px;
    line-height: 1.5;
}}

.flow-box {{
    background: #f0fdf4;
    border: 1px solid #bbf7d0;
    color: #064e3b;
    border-radius: 12px;
    padding: 8px;
    text-align: center;
    margin-bottom: 8px;
    font-weight: 800;
}}
</style>
""", unsafe_allow_html=True)


# =========================================================
# SIDEBAR CUSTOM
# =========================================================
def set_page(page_name):
    st.session_state.page = page_name


def sidebar_button(label, page_name, key):
    if st.session_state.page == page_name:
        st.sidebar.markdown(f'<div class="nav-active">{label}</div>', unsafe_allow_html=True)
    else:
        if st.sidebar.button(label, key=key, use_container_width=True):
            set_page(page_name)
            st.rerun()


st.sidebar.markdown("# 🌿 HyTBIONEX")
st.sidebar.markdown("Hybrid Transformer Pipeline")
st.sidebar.markdown('<div class="sidebar-line"></div>', unsafe_allow_html=True)

sidebar_button("🏠 Dashboard", "🏠 Dashboard", "nav_dashboard")

st.sidebar.markdown('<div class="sidebar-section-title">ANALISIS DATA</div>', unsafe_allow_html=True)
sidebar_button("🌿 Input Tanaman\n(Kata / Kalimat)", "🌿 Input Tanaman", "nav_input")
sidebar_button("📄 Upload Dokumen\n(PDF / Excel)", "📁 Upload Dokumen", "nav_upload")
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
st.markdown("""
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
""", unsafe_allow_html=True)


# =========================================================
# LOAD DATA
# =========================================================
df_data, dataset_status = load_dataset()

col_senyawa = find_col(df_data, ["Zat Bioaktif", "Senyawa Bioaktif", "Senyawa", "Compound", "Kandungan"])
col_khasiat = find_col(df_data, ["Khasiat", "Manfaat", "Benefit", "Khasiat/Efek Terapeutik", "Biological Activity"])

total_data = len(df_data)
total_senyawa = df_data[col_senyawa].astype(str).replace("", pd.NA).dropna().nunique() if col_senyawa else 0
total_khasiat = df_data[col_khasiat].astype(str).replace("", pd.NA).dropna().nunique() if col_khasiat else 0
total_relasi = total_data * 8 if total_data else 0


# =========================================================
# RENDER KOMPONEN
# =========================================================
def render_summary_card():
    st.markdown(f"""
    <div class="summary-card">
        <h3 style="color:#064e3b;margin-bottom:25px;">RINGKASAN DATA</h3>

        <div class="summary-item">
            <div class="summary-icon">🌿</div>
            <div>
                <div class="summary-label">Total Tanaman</div>
                <div class="summary-value">{total_data:,}</div>
            </div>
        </div>

        <div class="summary-item">
            <div class="summary-icon">🧪</div>
            <div>
                <div class="summary-label" style="color:#0284c7;">Total Senyawa</div>
                <div class="summary-value" style="color:#0284c7;">{total_senyawa:,}</div>
            </div>
        </div>

        <div class="summary-item">
            <div class="summary-icon">💗</div>
            <div>
                <div class="summary-label" style="color:#e11d48;">Total Khasiat</div>
                <div class="summary-value" style="color:#e11d48;">{total_khasiat:,}</div>
            </div>
        </div>

        <div class="summary-item">
            <div class="summary-icon">🔗</div>
            <div>
                <div class="summary-label" style="color:#7c3aed;">Relasi Triplet</div>
                <div class="summary-value" style="color:#7c3aed;">{total_relasi:,}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def render_single_input_area():
    st.markdown("""
    <div class="single-input-card">
        <h3>📝 Input Tanaman dan Dokumen</h3>
        <p>Masukkan nama tanaman atau kalimat, lalu unggah dokumen jika ada. Sistem akan mencocokkan informasi dengan dataset herbal.</p>
    </div>
    """, unsafe_allow_html=True)

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
                    dataset_status
                )

            st.success("Proses analisis selesai.")
            render_all_outputs(result, image_path, dataset_status, doc_status, match_status)


def render_quick_access():
    st.markdown('<div class="graph-box"><h3 style="color:#064e3b;">AKSES CEPAT ANALISIS</h3>', unsafe_allow_html=True)

    q1, q2, q3, q4 = st.columns(4)

    with q1:
        st.markdown("""
        <div class="quick-card">
            <div style="font-size:45px;">📖</div>
            <h4>Hasil Isolasi Entitas</h4>
            <p>Lihat entitas yang telah diidentifikasi</p>
            <div class="quick-btn">Lihat Data →</div>
        </div>
        """, unsafe_allow_html=True)

    with q2:
        st.markdown("""
        <div class="quick-card">
            <div style="font-size:45px;">🔗</div>
            <h4>Relation Extraction</h4>
            <p>Ekstraksi relasi antar entitas</p>
            <div class="quick-btn">Lihat Relasi →</div>
        </div>
        """, unsafe_allow_html=True)

    with q3:
        st.markdown("""
        <div class="quick-card">
            <div style="font-size:45px;">🧬</div>
            <h4>HerbKG 2.0 Explorer</h4>
            <p>Jelajahi graf pengetahuan tanaman herbal</p>
            <div class="quick-btn">Jelajah →</div>
        </div>
        """, unsafe_allow_html=True)

    with q4:
        st.markdown("""
        <div class="quick-card">
            <div style="font-size:45px;">📊</div>
            <h4>Statistik & Analitik</h4>
            <p>Visualisasi statistik dan analisis data</p>
            <div class="quick-btn">Lihat Statistik →</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("</div>", unsafe_allow_html=True)


def render_result_cards(result):
    st.markdown('<div class="section-title">📋 Hasil Ekstraksi Informasi Bioaktif</div>', unsafe_allow_html=True)

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
                st.markdown(f"""
                <div class="result-card">
                    <h4>{safe_text(title)}</h4>
                    <p>{safe_text(value)}</p>
                </div>
                """, unsafe_allow_html=True)
            idx += 1


def render_image_section(result, image_path):
    st.markdown('<div class="section-title">🖼️ Lampiran Gambar Tanaman</div>', unsafe_allow_html=True)

    if image_path:
        c1, c2 = st.columns([1, 2])
        with c1:
            st.image(image_path, caption=result["Nama Tanaman"], use_container_width=True)
        with c2:
            st.markdown(f"""
            <div class="lilac-card">
                <h3>🌿 Gambar Tanaman Terkoneksi Dataset</h3>
                <p><b>Nama Tanaman:</b> {safe_text(result["Nama Tanaman"])}</p>
                <p><b>Nama Latin:</b> {safe_text(result["Nama Latin"])}</p>
                <p><b>Catatan:</b> Gambar ini ditampilkan dari kolom <b>Gambar</b> pada dataset Excel atau nama file yang sesuai di folder <b>assets</b>.</p>
            </div>
            """, unsafe_allow_html=True)
    else:
        st.info("Gambar belum ditemukan. Pastikan Excel memiliki kolom Gambar, contoh isi: assets/jahe.jpg atau jahe.jpg.")


def render_relation_table(result):
    st.markdown('<div class="section-title">🔗 Relation Extraction</div>', unsafe_allow_html=True)

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
    st.markdown('<div class="section-title">🕸️ HerbKG 2.0 Explorer</div>', unsafe_allow_html=True)
    st.plotly_chart(make_kg_graph(result), use_container_width=True)


def render_descriptive_chart():
    st.markdown('<div class="section-title">📊 Grafik Analisis Deskriptif</div>', unsafe_allow_html=True)

    fig = make_descriptive_chart(df_data)

    if fig is None:
        st.info("Grafik belum dapat dibuat karena kolom nama tanaman tidak ditemukan.")
    else:
        st.plotly_chart(fig, use_container_width=True)


def render_all_outputs(result, image_path, dataset_status, doc_status, match_status):
    st.markdown(f"""
    <div class="lilac-card">
        <h3>📌 Status Sistem</h3>
        <p><b>Status Dataset:</b> {safe_text(dataset_status)}</p>
        <p><b>Status Dokumen:</b> {safe_text(doc_status)}</p>
        <p><b>Status Koneksi Entitas:</b> {safe_text(match_status)}</p>
    </div>
    """, unsafe_allow_html=True)

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

    st.markdown('<div class="summary-card"><h3 style="color:#064e3b;">PREVIEW KNOWLEDGE GRAPH</h3>', unsafe_allow_html=True)
    st.plotly_chart(make_kg_graph(sample), use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


def render_downstream_page():
    st.markdown('<div class="section-title">📦 Aplikasi Downstream</div>', unsafe_allow_html=True)

    st.markdown("""
    <div class="graph-box">
        <p style="color:#334155;font-size:16px;">
        Aplikasi downstream memanfaatkan hasil ekstraksi entitas dan relasi dari HerbKG 2.0
        untuk analisis deskriptif, query graf berbasis bukti, analisis kemiripan, dan rekomendasi herbal.
        </p>
    </div>
    """, unsafe_allow_html=True)

    d1, d2, d3, d4 = st.columns(4)

    with d1:
        st.markdown("""
        <div class="downstream-card">
            <h4>1. Analisis Deskriptif</h4>
            <p>Menampilkan ringkasan statistik entitas, relasi, dan distribusi data herbal.</p>
            <div class="flow-box">Tanaman</div>
            <div class="flow-box">Senyawa Bioaktif</div>
            <div class="flow-box">Khasiat</div>
            <div class="flow-box">Sumber Data</div>
        </div>
        """, unsafe_allow_html=True)

    with d2:
        st.markdown("""
        <div class="downstream-card">
            <h4>2. Query Graf Berbasis Bukti</h4>
            <p>Menelusuri hubungan tanaman, senyawa, khasiat, dan sumber literatur.</p>
            <div class="flow-box">Tanaman → Senyawa</div>
            <div class="flow-box">Senyawa → Khasiat</div>
            <div class="flow-box">Khasiat → Bukti Literatur</div>
            <div class="flow-box">Output: Jalur Relasi</div>
        </div>
        """, unsafe_allow_html=True)

    with d3:
        st.markdown("""
        <div class="downstream-card">
            <h4>3. Analisis Kemiripan</h4>
            <p>Menemukan tanaman yang mirip berdasarkan senyawa dan khasiat.</p>
            <div class="flow-box">Kelor → 0,56</div>
            <div class="flow-box">Sirih → 0,41</div>
            <div class="flow-box">Jahe → 0,39</div>
            <div class="flow-box">Kayu Manis → 0,28</div>
        </div>
        """, unsafe_allow_html=True)

    with d4:
        st.markdown("""
        <div class="downstream-card">
            <h4>4. Rekomendasi Herbal</h4>
            <p>Memberikan rekomendasi tanaman herbal berbasis khasiat dan relasi graf.</p>
            <div class="flow-box">Keluhan / Penyakit</div>
            <div class="flow-box">Graph Search</div>
            <div class="flow-box">Tanaman Terkait</div>
            <div class="flow-box">Peringkat Rekomendasi</div>
        </div>
        """, unsafe_allow_html=True)

    render_descriptive_chart()


# =========================================================
# HALAMAN
# =========================================================
if menu == "🏠 Dashboard":
    st.markdown('<div class="dashboard-shell">', unsafe_allow_html=True)

    top_left, top_right = st.columns([3.2, 1])

    with top_left:
        st.markdown("""
        <div class="hero-banner">
            <h2>Selamat datang di HyTBIONEX</h2>
            <p>
            Platform cerdas untuk isolasi informasi bioaktif dari tanaman herbal Indonesia
            menggunakan pendekatan Hybrid Transformer serta integrasi HerbKG 2.0.
            </p>
        </div>
        """, unsafe_allow_html=True)

    with top_right:
        render_summary_card()

    render_single_input_area()

    down_left, down_right = st.columns([2, 1])

    with down_left:
        render_quick_access()

    with down_right:
        render_preview_kg()

    f1, f2, f3 = st.columns(3)
    f1.markdown("<div class='footer-status'>⚙️ Model Aktif: Hybrid Transformer</div>", unsafe_allow_html=True)
    f2.markdown("<div class='footer-status'>🧬 Pipeline: NED → BIE → RE → HerbKG 2.0</div>", unsafe_allow_html=True)
    f3.markdown("<div class='footer-status'>🛡️ Status: Sistem siap digunakan</div>", unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)


elif menu in ["🌿 Input Tanaman", "📁 Upload Dokumen"]:
    render_single_input_area()


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
    st.markdown('<div class="section-title">📋 Cuplikan Dataset</div>', unsafe_allow_html=True)
    st.dataframe(df_data.head(30), use_container_width=True)


elif menu == "⚙️ Pengaturan":
    st.markdown("""
    <div class="lilac-card">
        <h2>⚙️ Pengaturan Sistem</h2>
        <p>Halaman ini disiapkan untuk pengaturan model, dataset, pipeline, dan tampilan sistem.</p>
    </div>
    """, unsafe_allow_html=True)


elif menu == "ℹ️ Tentang Aplikasi":
    st.markdown("""
    <div class="lilac-card">
        <h2>ℹ️ Tentang HyTBIONEX</h2>
        <p>
        HyTBIONEX adalah prototipe sistem ekstraksi informasi bioaktif tanaman herbal Indonesia
        berbasis pipeline Hybrid Transformer, ekstraksi entitas, ekstraksi relasi,
        visualisasi HerbKG 2.0, serta analisis deskriptif.
        </p>
        <p><b>Researcher:</b> Nazwita</p>
    </div>
    """, unsafe_allow_html=True)
