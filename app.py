import os
import re
import html
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

if "input_text" not in st.session_state:
    st.session_state.input_text = ""

if "file_key_suffix" not in st.session_state:
    st.session_state.file_key_suffix = 0


def clear_input_text():
    st.session_state.input_text = ""


def clear_uploaded_file():
    st.session_state.file_key_suffix += 1
    st.session_state.input_text = ""


def clear_text_when_upload():
    st.session_state.input_text = ""


# =====================================================
# CSS
# =====================================================
st.markdown("""
<style>
/* BACKGROUND UTAMA */
.stApp {
    background: linear-gradient(135deg, #e8f8ec 0%, #f3e8ff 100%);
}

/* SIDEBAR HIJAU */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #064e3b 0%, #047857 55%, #059669 100%) !important;
}

[data-testid="stSidebar"] * {
    color: #ffedd5 !important;
}

[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3 {
    color: #fed7aa !important;
    font-weight: 900 !important;
}

/* MENU RADIO SIDEBAR */
[data-testid="stSidebar"] [role="radiogroup"] label {
    background: rgba(255, 237, 213, 0.12) !important;
    padding: 13px 15px !important;
    border-radius: 14px !important;
    margin-bottom: 9px !important;
    border: 1px solid rgba(255, 237, 213, 0.18) !important;
}

[data-testid="stSidebar"] [role="radiogroup"] label:hover {
    background: rgba(249, 115, 22, 0.30) !important;
    transform: translateX(4px);
}

/* JUDUL UTAMA */
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
    border: 3px dashed #a855f7;
    padding: 22px;
    border-radius: 20px;
    margin-top: 18px;
    color: #111111;
}

/* INPUT TEXT AREA JADI LILAC */
textarea {
    background-color: #f3e8ff !important;
    color: #111111 !important;
    border: 2px solid #c084fc !important;
    border-radius: 16px !important;
}

textarea::placeholder {
    color: #6b4b84 !important;
}

/* FILE UPLOADER JADI LILAC */
[data-testid="stFileUploader"] section {
    background-color: #f3e8ff !important;
    border: 2px dashed #a855f7 !important;
    border-radius: 16px !important;
    color: #111111 !important;
}

[data-testid="stFileUploader"] section * {
    color: #111111 !important;
}

[data-testid="stFileUploader"] button {
    background: #a855f7 !important;
    color: white !important;
    border-radius: 12px !important;
    border: none !important;
    font-weight: 800 !important;
}

/* SEMUA TOMBOL TIDAK HITAM */
.stButton > button {
    background: linear-gradient(135deg, #f97316, #fb923c) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    font-weight: 900 !important;
    padding: 0.75rem 1rem !important;
}

/* TOMBOL X */
.x-button button {
    background: #a855f7 !important;
    color: white !important;
    border-radius: 50% !important;
    width: 42px !important;
    height: 42px !important;
    padding: 0 !important;
    font-size: 22px !important;
    border: 2px solid white !important;
}

.small-note {
    font-size: 15px;
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
    nama_tanaman = clean_text(get_col(row, [
        "Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"
    ]))

    nama_latin = clean_text(get_col(row, [
        "Nama Latin", "Nama_Latin", "Latin"
    ]))

    nama_lokal = clean_text(get_col(row, [
        "Nama Lokal/Daerah", "Nama Lokal", "Nama_Daerah",
        "Bahasa_Daerah", "Bahasa Daerah"
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


# =====================================================
# OUTPUT
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


def shorten_label(text, max_len=18):
    text = str(text)
    if text.lower() in ["", "nan", "none", "belum terdeteksi"]:
        return "Belum terdeteksi"
    if len(text) <= max_len:
        return text
    return text[:max_len] + "..."


def wrap_label(text, width=16):
    text = shorten_label(text, 32)
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

    st.markdown("## 🕸️ Enhanced Herb Knowledge Graph 2.0")

    G = nx.DiGraph()
    center = wrap_label(tanaman, 14)

    nodes = {
        center: "plant",
        wrap_label(latin, 18): "latin",
        wrap_label(bagian, 16): "part",
        wrap_label(zat, 16): "compound",
        wrap_label(khasiat, 16): "effect",
        wrap_label(olah, 16): "processing",
        wrap_label(dosis, 16): "dose",
        wrap_label(sumber, 22): "source"
    }

    for node, node_type in nodes.items():
        G.add_node(node, node_type=node_type)

    latin_node = wrap_label(latin, 18)
    bagian_node = wrap_label(bagian, 16)
    zat_node = wrap_label(zat, 16)
    khasiat_node = wrap_label(khasiat, 16)
    olah_node = wrap_label(olah, 16)
    dosis_node = wrap_label(dosis, 16)
    sumber_node = wrap_label(sumber, 22)

    edges = [
        (center, latin_node, "has_latin_name"),
        (center, bagian_node, "uses_part"),
        (center, zat_node, "contains"),
        (center, khasiat_node, "has_effect"),
        (center, olah_node, "processed_by"),
        (center, dosis_node, "has_dosage"),
        (center, sumber_node, "sourced_from"),
        (zat_node, khasiat_node, "contributes_to"),
        (bagian_node, zat_node, "contains_compound"),
    ]

    for source, target, relation in edges:
        G.add_edge(source, target, label=relation)

    pos = {
        center: (0.0, 0.2),
        latin_node: (-2.5, 1.4),
        bagian_node: (2.5, 1.4),
        zat_node: (-2.7, -0.3),
        khasiat_node: (0.0, -1.5),
        olah_node: (2.7, -0.3),
        dosis_node: (-2.4, -1.9),
        sumber_node: (2.3, -1.9),
    }

    node_colors = []
    node_edge_colors = []
    node_sizes = []

    for node in G.nodes():
        node_type = G.nodes[node]["node_type"]

        if node_type == "plant":
            node_colors.append("#ff5a5f")
            node_edge_colors.append("#ff2028")
            node_sizes.append(3600)
        elif node_type in ["compound", "dose"]:
            node_colors.append("#ffd6a5")
            node_edge_colors.append("#f97316")
            node_sizes.append(3000)
        elif node_type in ["effect", "latin", "source"]:
            node_colors.append("#f3c4fb")
            node_edge_colors.append("#a855f7")
            node_sizes.append(3300)
        else:
            node_colors.append("#70d99d")
            node_edge_colors.append("#00a85a")
            node_sizes.append(3000)

    fig, ax = plt.subplots(figsize=(12, 8))
    fig.patch.set_facecolor("#fbf7ff")
    ax.set_facecolor("#fbf7ff")

    nx.draw_networkx_edges(
        G,
        pos,
        ax=ax,
        arrows=True,
        arrowsize=22,
        arrowstyle="-|>",
        width=2.2,
        edge_color="#9ca3af",
        connectionstyle="arc3,rad=0.10"
    )

    nx.draw_networkx_nodes(
        G,
        pos,
        ax=ax,
        node_color=node_colors,
        edgecolors=node_edge_colors,
        linewidths=3,
        node_size=node_sizes
    )

    nx.draw_networkx_labels(
        G,
        pos,
        ax=ax,
        font_size=11,
        font_weight="bold",
        font_color="#111111"
    )

    edge_labels = nx.get_edge_attributes(G, "label")

    nx.draw_networkx_edge_labels(
        G,
        pos,
        edge_labels=edge_labels,
        ax=ax,
        font_size=9,
        font_color="#111111",
        rotate=True,
        label_pos=0.55,
        bbox=dict(
            boxstyle="round,pad=0.20",
            fc="#fbf7ff",
            ec="none",
            alpha=0.85
        )
    )

    ax.set_title(
        "Enhanced Herb Knowledge Graph 2.0",
        fontsize=18,
        fontweight="bold",
        color="#0b7a45",
        pad=20
    )

    ax.axis("off")
    plt.tight_layout()
    st.pyplot(fig, use_container_width=True)


# =====================================================
# SIDEBAR AKTIF
# =====================================================
st.sidebar.markdown("# 🌱 HyTBIONEX")
st.sidebar.markdown("---")

menu = st.sidebar.radio(
    "Pilih Menu",
    [
        "🏠 Dashboard Utama",
        "🌿 Input Data Tanaman",
        "📁 Upload Dokumen",
        "📋 Hasil Ekstraksi",
        "🔗 Relation Extraction",
        "🕸️ HerbKG 2.0",
        "📊 Descriptive Analytics",
        "🔎 Evidence-Based Graph Query",
        "🧬 Similarity Analysis",
        "💊 Herbal Recommendation"
    ],
    label_visibility="collapsed"
)

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

dataset_df, dataset_status = load_dataset()

# =====================================================
# MENU DASHBOARD
# =====================================================
if menu == "🏠 Dashboard Utama":
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

    st.markdown(
        f"""
        <div class="lilac-card">
            <h3>📌 Status Dataset</h3>
            <p>{safe(dataset_status)}</p>
        </div>
        """,
        unsafe_allow_html=True
    )

    c1, c2, c3 = st.columns(3)
    c1.metric("🌱 Total Data", len(dataset_df))
    c2.metric("📁 Dataset", "Terkoneksi" if not dataset_df.empty else "Belum")
    c3.metric("🕸️ HerbKG", "Aktif")


# =====================================================
# MENU INPUT / UPLOAD
# =====================================================
elif menu in ["🌿 Input Data Tanaman", "📁 Upload Dokumen"]:
    file_widget_key = f"uploaded_doc_{st.session_state.file_key_suffix}"
    dokumen_sebelumnya = st.session_state.get(file_widget_key, None)

    col1, col2 = st.columns(2)

    with col1:
        head1, x1 = st.columns([10, 1])
        with head1:
            st.markdown("""
            <div class="green-card">
                <h2>🌿 1. Input Data Tanaman</h2>
                <p>Masukkan nama tanaman atau kalimat artikel herbal.</p>
            </div>
            """, unsafe_allow_html=True)
        with x1:
            st.markdown('<div class="x-button">', unsafe_allow_html=True)
            st.button("×", key="clear_input", on_click=clear_input_text)
            st.markdown('</div>', unsafe_allow_html=True)

        if dokumen_sebelumnya is not None:
            st.markdown("""
            <div class="lilac-card">
                <h3>📁 Mode Dokumen Aktif</h3>
                <p>Input teks disembunyikan karena Ibu sedang menggunakan upload dokumen.
                Klik tombol X pada panel upload jika ingin kembali ke input teks.</p>
            </div>
            """, unsafe_allow_html=True)
            teks = ""
        else:
            teks = st.text_area(
                "Input Data Tanaman",
                placeholder="Contoh: Kelor, Sirih, Jahe, Kunyit, Kayu Manis, atau kalimat artikel herbal",
                height=180,
                key="input_text"
            )

    with col2:
        head2, x2 = st.columns([10, 1])
        with head2:
            st.markdown("""
            <div class="green-card">
                <h2>📁 2. Upload Dokumen Artikel / Dataset</h2>
                <p>Upload PDF, TXT, CSV, atau Excel.</p>
            </div>
            """, unsafe_allow_html=True)
        with x2:
            st.markdown('<div class="x-button">', unsafe_allow_html=True)
            st.button("×", key="clear_file", on_click=clear_uploaded_file)
            st.markdown('</div>', unsafe_allow_html=True)

        dokumen = st.file_uploader(
            "Upload Dokumen",
            type=["pdf", "txt", "csv", "xlsx", "xls"],
            key=file_widget_key,
            on_change=clear_text_when_upload
        )

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

        status.text("Membangun hasil ekstraksi... 100%")
        progress.progress(100)

        st.session_state.last_result = result
        st.session_state.last_dataset_status = dataset_status
        st.session_state.last_doc_status = doc_status
        st.session_state.last_match_status = match_status

        st.success("Proses ekstraksi selesai.")

        render_results(result, dataset_status, doc_status, match_status)
        render_relation(result)
        render_kg(result)


# =====================================================
# MENU HASIL EKSTRAKSI
# =====================================================
elif menu == "📋 Hasil Ekstraksi":
    if st.session_state.last_result:
        render_results(
            st.session_state.last_result,
            st.session_state.last_dataset_status,
            st.session_state.last_doc_status,
            st.session_state.last_match_status
        )
    else:
        st.warning("Belum ada hasil ekstraksi. Buka menu Input Data Tanaman, lalu klik PROSES EKSTRAKSI.")


elif menu == "🔗 Relation Extraction":
    if st.session_state.last_result:
        render_relation(st.session_state.last_result)
    else:
        st.warning("Belum ada hasil relasi. Jalankan proses ekstraksi terlebih dahulu.")


elif menu == "🕸️ HerbKG 2.0":
    if st.session_state.last_result:
        render_kg(st.session_state.last_result)
    else:
        st.warning("Belum ada HerbKG. Jalankan proses ekstraksi terlebih dahulu.")


elif menu == "📊 Descriptive Analytics":
    st.markdown("""
    <div class="lilac-card">
        <h2>📊 Descriptive Analytics</h2>
        <p>Menu ini digunakan untuk menampilkan ringkasan statistik dataset herbal.</p>
    </div>
    """, unsafe_allow_html=True)

    st.dataframe(dataset_df.head(20), use_container_width=True)


elif menu == "🔎 Evidence-Based Graph Query":
    st.markdown("""
    <div class="lilac-card">
        <h2>🔎 Evidence-Based Graph Query</h2>
        <p>Menu ini digunakan untuk menelusuri bukti relasi antara tanaman, senyawa bioaktif,
        khasiat, bagian tanaman, dosis, cara pengolahan, dan sumber literatur.</p>
    </div>
    """, unsafe_allow_html=True)


elif menu == "🧬 Similarity Analysis":
    st.markdown("""
    <div class="lilac-card">
        <h2>🧬 Similarity Analysis</h2>
        <p>Menu ini digunakan untuk menganalisis kemiripan tanaman herbal berdasarkan
        senyawa bioaktif, khasiat, bagian tanaman, dan pola relasi.</p>
    </div>
    """, unsafe_allow_html=True)


elif menu == "💊 Herbal Recommendation":
    st.markdown("""
    <div class="lilac-card">
        <h2>💊 Herbal Recommendation</h2>
        <p>Menu ini digunakan untuk rekomendasi tanaman herbal berdasarkan khasiat,
        senyawa bioaktif, atau kategori penyakit.</p>
    </div>
    """, unsafe_allow_html=True)
