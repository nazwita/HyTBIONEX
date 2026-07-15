import os
import re
import html
import base64
from pathlib import Path
from collections import Counter

import pandas as pd
import streamlit as st
import plotly.graph_objects as go
import plotly.express as px


# =========================================================
# KONFIGURASI APLIKASI
# =========================================================
APP_TITLE = "HyTBIONEX"
PREFERRED_DATASET = "Data set 20098+ Gambar.xlsx"
ASSET_DIR = "assets"

st.set_page_config(
    page_title=APP_TITLE,
    page_icon="🌿",
    layout="wide",
    initial_sidebar_state="expanded",
)


# =========================================================
# SESSION STATE
# =========================================================
DEFAULT_STATE = {
    "page": "🏠 Dashboard",
    "last_result": None,
    "last_image": None,
    "last_status": {},
    "dataset_file": None,
    "dashboard_text": "",
    "input_page_text": "",
    "upload_page_text": "",
}

for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value


# =========================================================
# FUNGSI UTILITAS
# =========================================================
def safe_text(value):
    if value is None:
        return ""
    return html.escape(str(value))


def clean_text(value):
    text = str(value).lower()
    text = re.sub(r"[^a-zA-Z0-9À-ÿ\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_colname(column):
    text = str(column).strip().lower()
    text = text.replace("_", " ").replace("/", " ")
    text = re.sub(r"[^a-zA-Z0-9À-ÿ\s]", " ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def slugify_filename(text):
    text = clean_text(text).replace(" ", "_")
    text = re.sub(r"_+", "_", text)
    return text.strip("_")


def short_label(text, max_len=28):
    text = str(text).strip()
    if text in {"", "nan", "None", "Belum terdeteksi"}:
        return "Belum terdeteksi"
    if len(text) <= max_len:
        return text
    return text[:max_len].rstrip() + "..."


def split_multi_value(value):
    if value is None:
        return []
    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "belum terdeteksi"}:
        return []
    parts = re.split(r"[,;/|]+", text)
    return [clean_text(p) for p in parts if clean_text(p)]


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


def discover_excel_files():
    files = []
    for item in os.listdir("."):
        if item.lower().endswith((".xlsx", ".xls")) and not item.startswith("~$"):
            files.append(item)
    return sorted(files)


def initialize_dataset_file():
    files = discover_excel_files()
    if st.session_state.dataset_file in files:
        return
    if PREFERRED_DATASET in files:
        st.session_state.dataset_file = PREFERRED_DATASET
    elif files:
        st.session_state.dataset_file = files[0]
    else:
        st.session_state.dataset_file = None


initialize_dataset_file()


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
    if df is None or df.empty:
        return None

    normalized_cols = {normalize_colname(c): c for c in df.columns}

    for candidate in candidates:
        candidate_norm = normalize_colname(candidate)
        if candidate_norm in normalized_cols:
            return normalized_cols[candidate_norm]

    for candidate in candidates:
        candidate_norm = normalize_colname(candidate)
        for normalized, original in normalized_cols.items():
            if candidate_norm in normalized or normalized in candidate_norm:
                return original

    return None


def value_from_row(row, column):
    if row is None or column is None:
        return "Belum terdeteksi"
    try:
        value = row.get(column, "")
    except Exception:
        return "Belum terdeteksi"

    if pd.isna(value) or str(value).strip() == "":
        return "Belum terdeteksi"
    return str(value).strip()


@st.cache_data(show_spinner=False)
def load_dataset(dataset_path):
    if not dataset_path or not os.path.exists(dataset_path):
        return pd.DataFrame(), "Dataset Excel belum ditemukan."

    try:
        sheets = pd.read_excel(dataset_path, sheet_name=None)
        frames = []

        for sheet_name, frame in sheets.items():
            if isinstance(frame, pd.DataFrame) and not frame.empty:
                frame = frame.copy().fillna("")
                frame["__sheet_name__"] = sheet_name
                frames.append(frame)

        if not frames:
            return pd.DataFrame(), f"Dataset kosong: {dataset_path}"

        data = pd.concat(frames, ignore_index=True).fillna("")
        data.columns = [str(c).strip() for c in data.columns]
        return data, (
            f"Dataset aktif: {dataset_path} | "
            f"{len(data):,} baris | {len(data.columns):,} kolom"
        )
    except Exception as error:
        return pd.DataFrame(), f"Gagal membaca dataset: {error}"


def read_uploaded_file(uploaded_file):
    if uploaded_file is None:
        return "", "Tidak ada dokumen yang diunggah."

    filename = uploaded_file.name.lower()

    try:
        if filename.endswith(".pdf"):
            from pypdf import PdfReader

            reader = PdfReader(uploaded_file)
            text = " ".join((page.extract_text() or "") for page in reader.pages)

            if not text.strip():
                return (
                    "",
                    f"PDF {uploaded_file.name} terbaca, tetapi teks kosong. "
                    "Kemungkinan PDF berupa hasil scan/gambar.",
                )

            return text, (
                f"PDF terbaca: {uploaded_file.name} | "
                f"{len(reader.pages)} halaman | {len(text):,} karakter"
            )

        if filename.endswith(".txt"):
            text = uploaded_file.read().decode("utf-8", errors="ignore")
            return text, (
                f"TXT terbaca: {uploaded_file.name} | {len(text):,} karakter"
            )

        if filename.endswith(".csv"):
            frame = pd.read_csv(uploaded_file).fillna("")
            text = " ".join(frame.astype(str).values.flatten())
            return text, (
                f"CSV terbaca: {uploaded_file.name} | {len(frame):,} baris"
            )

        if filename.endswith((".xlsx", ".xls")):
            sheets = pd.read_excel(uploaded_file, sheet_name=None)
            text_parts = []
            total_rows = 0

            for frame in sheets.values():
                frame = frame.fillna("")
                total_rows += len(frame)
                text_parts.extend(frame.astype(str).values.flatten())

            text = " ".join(map(str, text_parts))
            return text, (
                f"Excel terbaca: {uploaded_file.name} | "
                f"{len(sheets):,} sheet | {total_rows:,} baris"
            )

        return "", "Format dokumen belum didukung."

    except Exception as error:
        return "", f"Gagal membaca dokumen: {error}"


def get_column_map(df):
    return {
        "nama": find_col(df, [
            "Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama Herbal", "Nama"
        ]),
        "latin": find_col(df, [
            "Nama Latin", "Nama_Latin", "Latin", "Scientific Name"
        ]),
        "lokal": find_col(df, [
            "Nama Lokal/Daerah", "Nama Lokal", "Nama Daerah",
            "Bahasa Daerah", "Bahasa_Daerah", "Sinonim"
        ]),
        "bagian": find_col(df, [
            "Bagian Tanaman", "Bagian Digunakan", "Bagian_Digunakan",
            "Bagian yang Digunakan", "Bagian"
        ]),
        "senyawa": find_col(df, [
            "Zat Bioaktif", "Senyawa Bioaktif", "Senyawa_Bioaktif",
            "Compound", "Senyawa", "Kandungan", "Kandungan Kimia",
            "Komposisi/Kandungan Kimia"
        ]),
        "khasiat": find_col(df, [
            "Khasiat/Efek Terapeutik", "Khasiat", "Manfaat",
            "Benefit", "Biological Activity", "Biological_Activity",
            "Efek Terapeutik"
        ]),
        "pengolahan": find_col(df, [
            "Cara Pengolahan", "Cara_Pengolahan", "Pengolahan",
            "Cara Pemakaian", "Preparation"
        ]),
        "dosis": find_col(df, [
            "Komposisi/Dosis", "Komposisi /Dosis", "Dosis",
            "Komposisi", "Dose"
        ]),
        "sumber": find_col(df, [
            "Sumber Data", "Sumber_Data", "Sumber", "Referensi",
            "Source"
        ]),
        "gambar": find_col(df, [
            "Gambar", "Image", "Foto", "File Gambar",
            "Path Gambar", "Nama File Gambar"
        ]),
        "penyakit": find_col(df, [
            "Kategori Penyakit", "Penyakit", "Disease",
            "Nama Penyakit"
        ]),
        "benefit": find_col(df, [
            "Benefit", "Manfaat", "Deskripsi Manfaat"
        ]),
    }


def score_row(row, search_text, columns):
    score = 0
    search_text = clean_text(search_text)

    fields = [
        ("nama", 200, 30, 3),
        ("latin", 180, 24, 4),
        ("lokal", 120, 20, 3),
    ]

    for field, phrase_score, token_score, min_token_length in fields:
        column = columns.get(field)
        value = clean_text(value_from_row(row, column))

        if not value or value == "belum terdeteksi":
            continue

        values = split_multi_value(value) if field == "lokal" else [value]

        for item in values:
            if item and item in search_text:
                score += phrase_score

            for token in item.split():
                if len(token) >= min_token_length and token in search_text:
                    score += token_score

    return score


def find_best_match(df, search_text):
    if df.empty:
        return None, "Dataset belum terbaca.", 0

    search_text = clean_text(search_text)
    if not search_text:
        return None, "Input tanaman dan dokumen masih kosong.", 0

    columns = get_column_map(df)
    best_row = None
    best_score = 0

    for _, row in df.iterrows():
        current_score = score_row(row, search_text, columns)

        if current_score > best_score:
            best_score = current_score
            best_row = row

    if best_row is not None and best_score > 0:
        nama = value_from_row(best_row, columns["nama"])
        latin = value_from_row(best_row, columns["latin"])
        status = (
            f"Entitas cocok: {nama} / {latin} | Skor kecocokan: {best_score}"
        )
        return best_row, status, best_score

    return None, "Tidak ditemukan kecocokan tanaman pada dataset.", 0


def extract_result(row, input_text, df):
    columns = get_column_map(df)

    if row is None:
        return {
            "Nama Tanaman": input_text.strip() or "Belum terdeteksi",
            "Nama Latin": "Belum terdeteksi",
            "Nama Lokal/Daerah": "Belum terdeteksi",
            "Bagian Tanaman": "Belum terdeteksi",
            "Zat Bioaktif": "Belum terdeteksi",
            "Khasiat/Efek Terapeutik": "Belum terdeteksi",
            "Cara Pengolahan": "Belum terdeteksi",
            "Komposisi/Dosis": "Belum terdeteksi",
            "Sumber Data": "Belum terdeteksi",
            "Kategori Penyakit": "Belum terdeteksi",
            "Gambar": "Belum terdeteksi",
        }

    return {
        "Nama Tanaman": value_from_row(row, columns["nama"]),
        "Nama Latin": value_from_row(row, columns["latin"]),
        "Nama Lokal/Daerah": value_from_row(row, columns["lokal"]),
        "Bagian Tanaman": value_from_row(row, columns["bagian"]),
        "Zat Bioaktif": value_from_row(row, columns["senyawa"]),
        "Khasiat/Efek Terapeutik": value_from_row(row, columns["khasiat"]),
        "Cara Pengolahan": value_from_row(row, columns["pengolahan"]),
        "Komposisi/Dosis": value_from_row(row, columns["dosis"]),
        "Sumber Data": value_from_row(row, columns["sumber"]),
        "Kategori Penyakit": value_from_row(row, columns["penyakit"]),
        "Gambar": value_from_row(row, columns["gambar"]),
    }


def find_plant_image(result):
    image_value = result.get("Gambar", "")
    plant_name = result.get("Nama Tanaman", "")
    candidates = []

    if image_value and image_value != "Belum terdeteksi":
        candidates.extend([
            image_value,
            os.path.join(ASSET_DIR, image_value),
            os.path.join("gambar", image_value),
            os.path.join("images", image_value),
        ])

    slug = slugify_filename(plant_name)
    if slug:
        for extension in ["jpg", "jpeg", "png", "webp"]:
            candidates.extend([
                os.path.join(ASSET_DIR, f"{slug}.{extension}"),
                os.path.join("gambar", f"{slug}.{extension}"),
                os.path.join("images", f"{slug}.{extension}"),
                f"{slug}.{extension}",
            ])

    for path in candidates:
        if path and os.path.exists(path):
            return path

    return None


def run_extraction(text_input, uploaded_file, df, dataset_status):
    document_text, document_status = read_uploaded_file(uploaded_file)
    combined_text = f"{text_input} {document_text}".strip()

    row, match_status, score = find_best_match(df, combined_text)
    result = extract_result(row, text_input, df)
    image_path = find_plant_image(result)

    st.session_state.last_result = result
    st.session_state.last_image = image_path
    st.session_state.last_status = {
        "dataset": dataset_status,
        "document": document_status,
        "match": match_status,
        "score": score,
    }

    return result, image_path, document_status, match_status


# =========================================================
# FUNGSI NAVIGASI
# =========================================================
def set_page(page_name):
    st.session_state.page = page_name


def clear_analysis(prefix):
    text_key = f"{prefix}_text"
    upload_key = f"{prefix}_upload"

    if text_key in st.session_state:
        st.session_state[text_key] = ""
    if upload_key in st.session_state:
        st.session_state.pop(upload_key, None)

    st.session_state.last_result = None
    st.session_state.last_image = None
    st.session_state.last_status = {}


def sidebar_nav_button(label, page_name, key):
    if st.session_state.page == page_name:
        st.sidebar.markdown(
            f'<div class="nav-active">{safe_text(label)}</div>',
            unsafe_allow_html=True,
        )
    else:
        st.sidebar.button(
            label,
            key=key,
            use_container_width=True,
            on_click=set_page,
            args=(page_name,),
        )


# =========================================================
# VISUALISASI DATA
# =========================================================
def make_top_value_chart(df, column, title, x_title, top_n=10):
    if df.empty or column is None:
        return None

    series = (
        df[column]
        .astype(str)
        .str.strip()
        .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        .dropna()
    )

    if series.empty:
        return None

    chart_df = series.value_counts().head(top_n).reset_index()
    chart_df.columns = ["Kategori", "Jumlah"]

    fig = px.bar(
        chart_df,
        x="Jumlah",
        y="Kategori",
        orientation="h",
        text="Jumlah",
        title=title,
    )
    fig.update_traces(textposition="outside")
    fig.update_yaxes(autorange="reversed")
    fig.update_layout(
        height=430,
        plot_bgcolor="white",
        paper_bgcolor="white",
        xaxis_title=x_title,
        yaxis_title="",
        margin=dict(l=20, r=30, t=65, b=20),
        coloraxis_showscale=False,
    )
    return fig


def make_kg_graph(result):
    values = {
        "tanaman": result.get("Nama Tanaman", "Belum terdeteksi"),
        "latin": result.get("Nama Latin", "Belum terdeteksi"),
        "lokal": result.get("Nama Lokal/Daerah", "Belum terdeteksi"),
        "bagian": result.get("Bagian Tanaman", "Belum terdeteksi"),
        "senyawa": result.get("Zat Bioaktif", "Belum terdeteksi"),
        "khasiat": result.get(
            "Khasiat/Efek Terapeutik", "Belum terdeteksi"
        ),
        "pengolahan": result.get("Cara Pengolahan", "Belum terdeteksi"),
        "dosis": result.get("Komposisi/Dosis", "Belum terdeteksi"),
        "sumber": result.get("Sumber Data", "Belum terdeteksi"),
    }

    nodes = [
        {
            "id": "tanaman", "x": 0, "y": 0,
            "label": short_label(values["tanaman"], 18),
            "full": values["tanaman"],
            "color": "#047857", "border": "#065f46", "size": 88,
        },
        {
            "id": "latin", "x": -2.8, "y": 1.55,
            "label": short_label(values["latin"], 25),
            "full": values["latin"],
            "color": "#e9d5ff", "border": "#9333ea", "size": 66,
        },
        {
            "id": "lokal", "x": 0, "y": 2.25,
            "label": short_label(values["lokal"], 23),
            "full": values["lokal"],
            "color": "#bfdbfe", "border": "#2563eb", "size": 66,
        },
        {
            "id": "bagian", "x": 2.8, "y": 1.55,
            "label": short_label(values["bagian"], 20),
            "full": values["bagian"],
            "color": "#bbf7d0", "border": "#047857", "size": 66,
        },
        {
            "id": "senyawa", "x": -3.15, "y": 0,
            "label": short_label(values["senyawa"], 28),
            "full": values["senyawa"],
            "color": "#fed7aa", "border": "#f97316", "size": 74,
        },
        {
            "id": "pengolahan", "x": 3.15, "y": 0,
            "label": short_label(values["pengolahan"], 28),
            "full": values["pengolahan"],
            "color": "#bbf7d0", "border": "#047857", "size": 72,
        },
        {
            "id": "dosis", "x": -2.45, "y": -1.75,
            "label": short_label(values["dosis"], 25),
            "full": values["dosis"],
            "color": "#fde68a", "border": "#d97706", "size": 70,
        },
        {
            "id": "khasiat", "x": 0, "y": -2.25,
            "label": short_label(values["khasiat"], 27),
            "full": values["khasiat"],
            "color": "#fbcfe8", "border": "#ec4899", "size": 76,
        },
        {
            "id": "sumber", "x": 2.45, "y": -1.75,
            "label": short_label(values["sumber"], 28),
            "full": values["sumber"],
            "color": "#e9d5ff", "border": "#9333ea", "size": 70,
        },
    ]

    node_map = {node["id"]: node for node in nodes}

    edges = [
        ("tanaman", "latin", "nama latin"),
        ("tanaman", "lokal", "nama lokal/daerah"),
        ("tanaman", "bagian", "bagian digunakan"),
        ("tanaman", "senyawa", "mengandung"),
        ("tanaman", "pengolahan", "cara pengolahan"),
        ("tanaman", "dosis", "dosis/komposisi"),
        ("tanaman", "khasiat", "memiliki khasiat"),
        ("tanaman", "sumber", "sumber data"),
    ]

    label_shift = {
        "nama latin": (0, 12),
        "nama lokal/daerah": (55, 0),
        "bagian digunakan": (0, 12),
        "mengandung": (0, 15),
        "cara pengolahan": (0, 15),
        "dosis/komposisi": (-5, 0),
        "memiliki khasiat": (55, 0),
        "sumber data": (5, 0),
    }

    fig = go.Figure()

    for source, target, relation in edges:
        source_node = node_map[source]
        target_node = node_map[target]

        fig.add_trace(go.Scatter(
            x=[source_node["x"], target_node["x"]],
            y=[source_node["y"], target_node["y"]],
            mode="lines",
            line=dict(color="#94a3b8", width=2.5),
            hoverinfo="none",
            showlegend=False,
        ))

        x_shift, y_shift = label_shift.get(relation, (0, 0))

        fig.add_annotation(
            x=(source_node["x"] + target_node["x"]) / 2,
            y=(source_node["y"] + target_node["y"]) / 2,
            text=relation,
            showarrow=False,
            xshift=x_shift,
            yshift=y_shift,
            font=dict(size=12, color="#111827"),
            bgcolor="rgba(255,255,255,0.92)",
            bordercolor="#dbe3ec",
            borderwidth=1,
            borderpad=4,
        )

    for node in nodes:
        hover_text = (
            f"<b>{safe_text(node['label'])}</b>"
            f"<br>{safe_text(node['full'])}"
        )

        fig.add_trace(go.Scatter(
            x=[node["x"]],
            y=[node["y"]],
            mode="markers+text",
            marker=dict(
                size=node["size"],
                color=node["color"],
                line=dict(color=node["border"], width=4),
            ),
            text=[node["label"]],
            textposition="middle center",
            textfont=dict(
                size=13,
                color="#111111",
                family="Arial Black",
            ),
            hovertext=[hover_text],
            hoverinfo="text",
            showlegend=False,
            cliponaxis=False,
        ))

    fig.update_layout(
        height=720,
        plot_bgcolor="#fbf7ff",
        paper_bgcolor="#fbf7ff",
        margin=dict(l=20, r=20, t=30, b=30),
        xaxis=dict(visible=False, range=[-4.2, 4.2], fixedrange=True),
        yaxis=dict(visible=False, range=[-3.0, 2.9], fixedrange=True),
        dragmode=False,
        hoverlabel=dict(
            bgcolor="white",
            font_size=13,
            font_family="Arial",
        ),
    )

    return fig


def build_relation_dataframe(result):
    plant = result.get("Nama Tanaman", "Belum terdeteksi")
    return pd.DataFrame([
        [plant, "memiliki nama latin", result.get("Nama Latin", "")],
        [plant, "memiliki nama lokal/daerah", result.get("Nama Lokal/Daerah", "")],
        [plant, "menggunakan bagian tanaman", result.get("Bagian Tanaman", "")],
        [plant, "mengandung senyawa bioaktif", result.get("Zat Bioaktif", "")],
        [plant, "memiliki khasiat", result.get("Khasiat/Efek Terapeutik", "")],
        [plant, "diolah dengan cara", result.get("Cara Pengolahan", "")],
        [plant, "memiliki dosis/komposisi", result.get("Komposisi/Dosis", "")],
        [plant, "bersumber dari", result.get("Sumber Data", "")],
    ], columns=["Entitas Sumber", "Relasi", "Entitas Tujuan"])


# =========================================================
# ANALISIS DOWNSTREAM
# =========================================================
def unique_plant_names(df):
    columns = get_column_map(df)
    column = columns["nama"]

    if df.empty or column is None:
        return []

    values = (
        df[column].astype(str).str.strip()
        .replace({"": pd.NA, "nan": pd.NA})
        .dropna()
        .unique()
        .tolist()
    )
    return sorted(values)


def find_rows_by_plant(df, plant_name):
    columns = get_column_map(df)
    name_col = columns["nama"]

    if df.empty or name_col is None:
        return pd.DataFrame()

    mask = (
        df[name_col].astype(str).str.strip().str.lower()
        == str(plant_name).strip().lower()
    )
    return df.loc[mask].copy()


def merge_rows_to_result(rows, df):
    if rows.empty:
        return extract_result(None, "", df)

    columns = get_column_map(df)
    result = {}

    mapping = {
        "Nama Tanaman": "nama",
        "Nama Latin": "latin",
        "Nama Lokal/Daerah": "lokal",
        "Bagian Tanaman": "bagian",
        "Zat Bioaktif": "senyawa",
        "Khasiat/Efek Terapeutik": "khasiat",
        "Cara Pengolahan": "pengolahan",
        "Komposisi/Dosis": "dosis",
        "Sumber Data": "sumber",
        "Kategori Penyakit": "penyakit",
        "Gambar": "gambar",
    }

    for output_key, column_key in mapping.items():
        column = columns.get(column_key)

        if column is None:
            result[output_key] = "Belum terdeteksi"
            continue

        values = [
            str(value).strip()
            for value in rows[column].tolist()
            if str(value).strip() and str(value).strip().lower() != "nan"
        ]

        unique_values = list(dict.fromkeys(values))
        result[output_key] = (
            " | ".join(unique_values[:5])
            if unique_values
            else "Belum terdeteksi"
        )

    return result


def calculate_jaccard_similarity(base_set, candidate_set):
    if not base_set and not candidate_set:
        return 0.0
    union = base_set | candidate_set
    if not union:
        return 0.0
    return len(base_set & candidate_set) / len(union)


def calculate_similarity_table(df, selected_plant):
    columns = get_column_map(df)
    name_col = columns["nama"]
    compound_col = columns["senyawa"]
    effect_col = columns["khasiat"]

    if df.empty or name_col is None:
        return pd.DataFrame()

    selected_rows = find_rows_by_plant(df, selected_plant)
    if selected_rows.empty:
        return pd.DataFrame()

    selected_compounds = set()
    selected_effects = set()

    if compound_col:
        for value in selected_rows[compound_col]:
            selected_compounds.update(split_multi_value(value))

    if effect_col:
        for value in selected_rows[effect_col]:
            selected_effects.update(split_multi_value(value))

    results = []

    for plant in unique_plant_names(df):
        if plant == selected_plant:
            continue

        rows = find_rows_by_plant(df, plant)
        compounds = set()
        effects = set()

        if compound_col:
            for value in rows[compound_col]:
                compounds.update(split_multi_value(value))

        if effect_col:
            for value in rows[effect_col]:
                effects.update(split_multi_value(value))

        compound_similarity = calculate_jaccard_similarity(
            selected_compounds, compounds
        )
        effect_similarity = calculate_jaccard_similarity(
            selected_effects, effects
        )
        total_similarity = (
            0.55 * compound_similarity
            + 0.45 * effect_similarity
        )

        results.append({
            "Tanaman Pembanding": plant,
            "Kemiripan Senyawa": round(compound_similarity, 4),
            "Kemiripan Khasiat": round(effect_similarity, 4),
            "Skor Kemiripan": round(total_similarity, 4),
        })

    if not results:
        return pd.DataFrame()

    return (
        pd.DataFrame(results)
        .sort_values("Skor Kemiripan", ascending=False)
        .reset_index(drop=True)
    )


def search_recommendations(df, keyword):
    if df.empty or not keyword.strip():
        return pd.DataFrame()

    columns = get_column_map(df)
    searchable_columns = [
        columns["khasiat"],
        columns["benefit"],
        columns["penyakit"],
        columns["senyawa"],
    ]
    searchable_columns = [c for c in searchable_columns if c is not None]

    if not searchable_columns:
        return pd.DataFrame()

    keyword_clean = clean_text(keyword)
    scores = []

    for index, row in df.iterrows():
        score = 0
        evidence = []

        for column in searchable_columns:
            value = str(row.get(column, "")).strip()
            value_clean = clean_text(value)

            if keyword_clean and keyword_clean in value_clean:
                score += 3
                evidence.append(f"{column}: {value}")

            for token in keyword_clean.split():
                if len(token) >= 3 and token in value_clean:
                    score += 1

        if score > 0:
            scores.append({
                "_index": index,
                "Skor": score,
                "Bukti": " | ".join(evidence[:3]),
            })

    if not scores:
        return pd.DataFrame()

    score_df = pd.DataFrame(scores)
    matched = df.loc[score_df["_index"]].copy().reset_index(drop=True)
    matched["Skor Rekomendasi"] = score_df["Skor"].values
    matched["Bukti Kecocokan"] = score_df["Bukti"].values

    name_col = columns["nama"]
    latin_col = columns["latin"]
    effect_col = columns["khasiat"]
    compound_col = columns["senyawa"]
    part_col = columns["bagian"]
    source_col = columns["sumber"]

    output = pd.DataFrame({
        "Nama Tanaman": (
            matched[name_col] if name_col else "Belum terdeteksi"
        ),
        "Nama Latin": (
            matched[latin_col] if latin_col else "Belum terdeteksi"
        ),
        "Khasiat/Efek": (
            matched[effect_col] if effect_col else "Belum terdeteksi"
        ),
        "Zat Bioaktif": (
            matched[compound_col] if compound_col else "Belum terdeteksi"
        ),
        "Bagian": (
            matched[part_col] if part_col else "Belum terdeteksi"
        ),
        "Sumber": (
            matched[source_col] if source_col else "Belum terdeteksi"
        ),
        "Skor Rekomendasi": matched["Skor Rekomendasi"],
        "Bukti Kecocokan": matched["Bukti Kecocokan"],
    })

    return (
        output
        .sort_values("Skor Rekomendasi", ascending=False)
        .drop_duplicates(subset=["Nama Tanaman", "Khasiat/Efek"])
        .reset_index(drop=True)
    )


# =========================================================
# CSS
# =========================================================
hero_bg = find_optional_background()

if hero_bg:
    hero_background_css = (
        "linear-gradient(90deg, rgba(255,255,255,0.97), "
        "rgba(255,255,255,0.60)), "
        f'url("{hero_bg}")'
    )
else:
    hero_background_css = (
        "radial-gradient(circle at 70% 40%, "
        "rgba(187,247,208,0.90), transparent 28%), "
        "linear-gradient(135deg, #ffffff 0%, #ecfdf5 55%, #f8fafc 100%)"
    )

st.markdown(
    f"""<style>
.stApp {{
    background: #f5fbf7;
}}
.main .block-container {{
    padding-top: 1rem;
    max-width: 1500px;
}}
[data-testid="stSidebar"] {{
    background:
        radial-gradient(circle at bottom left, rgba(34,197,94,0.25), transparent 28%),
        radial-gradient(circle at top right, rgba(16,185,129,0.18), transparent 35%),
        linear-gradient(180deg, #021f16 0%, #043b2c 45%, #065f46 100%);
    border-right: 1px solid rgba(255,255,255,0.12);
}}
[data-testid="stSidebar"] * {{
    color: #f8fff8;
}}
[data-testid="stSidebar"] .stButton > button {{
    background: transparent;
    color: #f8fff8;
    border: none;
    box-shadow: none;
    text-align: left;
    justify-content: flex-start;
    padding: 0.72rem 0.8rem;
    border-radius: 0.8rem;
    font-weight: 650;
    margin-bottom: 0.2rem;
    width: 100%;
}}
[data-testid="stSidebar"] .stButton > button:hover {{
    background: rgba(255,255,255,0.12);
    color: #ffffff;
}}
.nav-active {{
    background: linear-gradient(90deg, rgba(34,197,94,0.38), rgba(16,185,129,0.20));
    border: 1px solid rgba(134,239,172,0.28);
    border-radius: 0.8rem;
    padding: 0.78rem 0.85rem;
    margin-bottom: 0.3rem;
    color: #ffffff;
    font-weight: 800;
}}
.sidebar-section-title {{
    color: #86efac;
    font-size: 0.75rem;
    font-weight: 900;
    letter-spacing: 0.08rem;
    margin-top: 1.25rem;
    margin-bottom: 0.45rem;
}}
.sidebar-line {{
    height: 1px;
    background: rgba(255,255,255,0.14);
    margin: 1rem 0;
}}
.top-header {{
    background: linear-gradient(135deg, #013220, #064e3b, #065f46);
    padding: 1.4rem 1.7rem;
    border-radius: 0 0 1.6rem 1.6rem;
    color: white;
    margin-bottom: 1.25rem;
    box-shadow: 0 16px 38px rgba(0,0,0,0.18);
}}
.top-header h1 {{
    margin: 0;
    color: white;
    font-size: 2.1rem;
    font-weight: 900;
}}
.top-header p {{
    margin: 0.35rem 0 0;
    color: #d1fae5;
}}
.hero-banner {{
    background: {hero_background_css};
    background-size: cover;
    background-position: center;
    padding: 2rem;
    border-radius: 1.35rem;
    min-height: 220px;
    box-shadow: inset 0 0 0 1px rgba(6,78,59,0.08);
}}
.hero-banner h2 {{
    color: #047857;
    font-size: 1.85rem;
    font-weight: 900;
    margin-bottom: 0.8rem;
}}
.hero-banner p {{
    color: #12372a;
    font-size: 1rem;
    line-height: 1.65;
    max-width: 700px;
}}
.section-title {{
    color: #064e3b;
    font-size: 1.55rem;
    font-weight: 900;
    margin-top: 1.25rem;
    margin-bottom: 0.75rem;
}}
[data-testid="stMetric"] {{
    background: white;
    border: 1px solid rgba(6,78,59,0.10);
    border-radius: 1rem;
    padding: 1rem;
    box-shadow: 0 8px 20px rgba(15,23,42,0.07);
}}
[data-testid="stMetricLabel"] {{
    color: #065f46;
    font-weight: 800;
}}
[data-testid="stMetricValue"] {{
    color: #047857;
    font-weight: 900;
}}
.result-card {{
    background: #ffffff;
    border-left: 0.42rem solid #047857;
    border-radius: 0.9rem;
    padding: 1rem;
    min-height: 118px;
    box-shadow: 0 8px 18px rgba(15,23,42,0.08);
    margin-bottom: 0.75rem;
}}
.result-card h4 {{
    color: #064e3b;
    font-weight: 900;
    margin: 0 0 0.45rem 0;
}}
.result-card p {{
    color: #0f172a;
    font-size: 0.98rem;
    line-height: 1.45;
    margin: 0;
}}
.info-box {{
    background: #f3e8ff;
    padding: 1rem;
    border-radius: 1rem;
    border: 1px solid #c084fc;
    color: #111827;
    margin-bottom: 1rem;
}}
.stButton > button {{
    background: linear-gradient(90deg, #047857, #059669);
    color: white;
    border: none;
    border-radius: 0.8rem;
    font-weight: 900;
}}
.stButton > button:hover {{
    background: linear-gradient(90deg, #065f46, #047857);
    color: white;
}}
[data-testid="stFileUploader"] section {{
    background: #fbfffb;
    border: 2px dashed #86efac;
    border-radius: 1rem;
}}
textarea {{
    background: white !important;
    color: #0f172a !important;
}}
.quick-card {{
    background: white;
    border-radius: 1rem;
    padding: 1rem;
    border: 1px solid rgba(6,78,59,0.10);
    box-shadow: 0 8px 20px rgba(15,23,42,0.07);
    min-height: 160px;
    text-align: center;
}}
.quick-card h4 {{
    color: #064e3b;
    margin: 0.4rem 0;
}}
.quick-card p {{
    color: #475569;
    font-size: 0.88rem;
}}
</style>""",
    unsafe_allow_html=True,
)


# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.markdown("# 🌿 HyTBIONEX")
st.sidebar.caption("Hybrid Transformer Pipeline")
st.sidebar.markdown('<div class="sidebar-line"></div>', unsafe_allow_html=True)

sidebar_nav_button("🏠 Dashboard", "🏠 Dashboard", "nav_dashboard")

st.sidebar.markdown(
    '<div class="sidebar-section-title">EKSTRAKSI DATA</div>',
    unsafe_allow_html=True,
)
sidebar_nav_button("🌿 Input Tanaman", "🌿 Input Tanaman", "nav_input")
sidebar_nav_button("📄 Upload Dokumen", "📄 Upload Dokumen", "nav_upload")
sidebar_nav_button(
    "📋 Hasil Isolasi Entitas",
    "📋 Hasil Isolasi Entitas",
    "nav_entity_result",
)
sidebar_nav_button(
    "🔗 Relation Extraction",
    "🔗 Relation Extraction",
    "nav_relation",
)
sidebar_nav_button(
    "🕸️ HerbKG 2.0 Explorer",
    "🕸️ HerbKG 2.0 Explorer",
    "nav_kg",
)

st.sidebar.markdown(
    '<div class="sidebar-section-title">APLIKASI DOWNSTREAM</div>',
    unsafe_allow_html=True,
)
sidebar_nav_button(
    "📦 Ringkasan Downstream",
    "📦 Ringkasan Downstream",
    "nav_downstream",
)
sidebar_nav_button(
    "📊 Analisis Deskriptif",
    "📊 Analisis Deskriptif",
    "nav_descriptive",
)
sidebar_nav_button(
    "🔎 Query Graf Berbasis Bukti",
    "🔎 Query Graf Berbasis Bukti",
    "nav_evidence",
)
sidebar_nav_button(
    "🧬 Analisis Kemiripan",
    "🧬 Analisis Kemiripan",
    "nav_similarity",
)
sidebar_nav_button(
    "💡 Rekomendasi Herbal",
    "💡 Rekomendasi Herbal",
    "nav_recommendation",
)

st.sidebar.markdown(
    '<div class="sidebar-section-title">MODEL DAN SISTEM</div>',
    unsafe_allow_html=True,
)
sidebar_nav_button(
    "📈 Statistik & Analitik",
    "📈 Statistik & Analitik",
    "nav_statistics",
)
sidebar_nav_button(
    "🧩 Training Model",
    "🧩 Training Model",
    "nav_training",
)
sidebar_nav_button(
    "⚙️ Pengaturan",
    "⚙️ Pengaturan",
    "nav_settings",
)
sidebar_nav_button(
    "ℹ️ Tentang Aplikasi",
    "ℹ️ Tentang Aplikasi",
    "nav_about",
)

st.sidebar.markdown('<div class="sidebar-line"></div>', unsafe_allow_html=True)
st.sidebar.caption("© 2026 HyTBIONEX • Nazwita, M.Kom.")


# =========================================================
# HEADER
# =========================================================
st.markdown(
    """<div class="top-header">
<h1>🌿 HyTBIONEX</h1>
<p>Sistem Ekstraksi Informasi Bioaktif Tanaman Herbal Indonesia dan HerbKG 2.0</p>
</div>""",
    unsafe_allow_html=True,
)


# =========================================================
# LOAD DATASET
# =========================================================
df_data, dataset_status = load_dataset(st.session_state.dataset_file)
columns = get_column_map(df_data)

total_rows = len(df_data)
total_plants = (
    df_data[columns["nama"]].astype(str).replace("", pd.NA).dropna().nunique()
    if columns["nama"] else 0
)
total_compounds = (
    df_data[columns["senyawa"]].astype(str).replace("", pd.NA).dropna().nunique()
    if columns["senyawa"] else 0
)
total_effects = (
    df_data[columns["khasiat"]].astype(str).replace("", pd.NA).dropna().nunique()
    if columns["khasiat"] else 0
)
total_relations = total_rows * 8 if total_rows else 0


# =========================================================
# KOMPONEN TAMPILAN
# =========================================================
def render_metrics():
    metric_cols = st.columns(5)
    metric_cols[0].metric("Total Baris", f"{total_rows:,}")
    metric_cols[1].metric("Total Tanaman", f"{total_plants:,}")
    metric_cols[2].metric("Total Senyawa", f"{total_compounds:,}")
    metric_cols[3].metric("Total Khasiat", f"{total_effects:,}")
    metric_cols[4].metric("Relasi Triplet", f"{total_relations:,}")


def render_status_box(dataset_text, document_text, match_text):
    st.markdown(
        f"""<div class="info-box">
<b>Status Dataset:</b> {safe_text(dataset_text)}<br>
<b>Status Dokumen:</b> {safe_text(document_text)}<br>
<b>Status Koneksi Entitas:</b> {safe_text(match_text)}
</div>""",
        unsafe_allow_html=True,
    )


def render_analysis_form(prefix, allow_text=True, allow_upload=True):
    st.markdown(
        '<div class="section-title">📝 Input dan Proses Ekstraksi</div>',
        unsafe_allow_html=True,
    )

    text_input = ""

    if allow_text and allow_upload:
        left, right = st.columns([1.1, 1])

        with left:
            text_input = st.text_area(
                "Nama tanaman atau kalimat",
                placeholder=(
                    "Contoh: Serai mengandung citronellal dan digunakan "
                    "sebagai antimikroba."
                ),
                height=145,
                key=f"{prefix}_text",
            )

        with right:
            uploaded_file = st.file_uploader(
                "Unggah dokumen",
                type=["pdf", "txt", "csv", "xlsx", "xls"],
                key=f"{prefix}_upload",
            )

    elif allow_text:
        text_input = st.text_area(
            "Nama tanaman atau kalimat",
            placeholder="Contoh: Serai atau Cymbopogon nardus",
            height=150,
            key=f"{prefix}_text",
        )
        uploaded_file = None

    else:
        uploaded_file = st.file_uploader(
            "Unggah dokumen",
            type=["pdf", "txt", "csv", "xlsx", "xls"],
            key=f"{prefix}_upload",
        )

    process_col, clear_col = st.columns([4, 1])

    with process_col:
        submitted = st.button(
            "🔍 Proses Ekstraksi",
            use_container_width=True,
            key=f"{prefix}_process",
        )

    with clear_col:
        st.button(
            "🗑️ Bersihkan",
            use_container_width=True,
            key=f"{prefix}_clear",
            on_click=clear_analysis,
            args=(prefix,),
        )

    if submitted:
        if not text_input.strip() and uploaded_file is None:
            st.warning(
                "Masukkan nama tanaman/kalimat atau unggah dokumen terlebih dahulu."
            )
            return

        with st.spinner("Sedang melakukan proses ekstraksi informasi bioaktif..."):
            result, image_path, document_status, match_status = run_extraction(
                text_input,
                uploaded_file,
                df_data,
                dataset_status,
            )

        st.success("Proses ekstraksi berhasil dilakukan.")
        render_all_outputs(
            result=result,
            image_path=image_path,
            dataset_text=dataset_status,
            document_text=document_status,
            match_text=match_status,
            key_prefix=f"{prefix}_output",
        )


def render_result_cards(result):
    st.markdown(
        '<div class="section-title">📋 Hasil Ekstraksi Informasi Bioaktif</div>',
        unsafe_allow_html=True,
    )

    cards = [
        ("🌿 Nama Tanaman", result.get("Nama Tanaman", "")),
        ("🔬 Nama Latin", result.get("Nama Latin", "")),
        ("🇮🇩 Nama Lokal/Daerah", result.get("Nama Lokal/Daerah", "")),
        ("🍃 Bagian Tanaman", result.get("Bagian Tanaman", "")),
        ("🧪 Zat Bioaktif", result.get("Zat Bioaktif", "")),
        (
            "💚 Khasiat/Efek Terapeutik",
            result.get("Khasiat/Efek Terapeutik", ""),
        ),
        ("☕ Cara Pengolahan", result.get("Cara Pengolahan", "")),
        ("⚖️ Komposisi/Dosis", result.get("Komposisi/Dosis", "")),
        ("📚 Sumber Data", result.get("Sumber Data", "")),
    ]

    for start in range(0, len(cards), 3):
        row = st.columns(3)
        for column, (title, value) in zip(row, cards[start:start + 3]):
            with column:
                st.markdown(
                    f"""<div class="result-card">
<h4>{safe_text(title)}</h4>
<p>{safe_text(value)}</p>
</div>""",
                    unsafe_allow_html=True,
                )


def render_image_section(result, image_path):
    st.markdown(
        '<div class="section-title">🖼️ Gambar Tanaman</div>',
        unsafe_allow_html=True,
    )

    if image_path:
        left, right = st.columns([1, 2])

        with left:
            st.image(
                image_path,
                caption=result.get("Nama Tanaman", ""),
                use_container_width=True,
            )

        with right:
            st.info(
                "Gambar ditemukan dari kolom Gambar pada dataset "
                "atau dari folder assets berdasarkan nama tanaman."
            )
            st.write("**Nama Tanaman:**", result.get("Nama Tanaman", ""))
            st.write("**Nama Latin:**", result.get("Nama Latin", ""))
    else:
        st.info(
            "Gambar belum ditemukan. Simpan gambar di folder assets, "
            "contoh: assets/serai.jpg, atau isi kolom Gambar pada dataset."
        )


def render_relation_table(result, key_prefix):
    st.markdown(
        '<div class="section-title">🔗 Relation Extraction</div>',
        unsafe_allow_html=True,
    )
    relation_df = build_relation_dataframe(result)
    st.dataframe(
        relation_df,
        use_container_width=True,
        hide_index=True,
    )
    st.download_button(
        "⬇️ Unduh Relasi CSV",
        data=relation_df.to_csv(index=False).encode("utf-8"),
        file_name="relation_extraction_hytbionex.csv",
        mime="text/csv",
        key=f"{key_prefix}_download_relation",
    )


def render_kg_section(result, chart_key):
    st.markdown(
        '<div class="section-title">🕸️ HerbKG 2.0 Explorer</div>',
        unsafe_allow_html=True,
    )
    st.plotly_chart(
        make_kg_graph(result),
        use_container_width=True,
        key=chart_key,
        config={
            "displayModeBar": True,
            "scrollZoom": False,
            "displaylogo": False,
        },
    )


def render_descriptive_charts(key_prefix):
    st.markdown(
        '<div class="section-title">📊 Analisis Deskriptif</div>',
        unsafe_allow_html=True,
    )

    chart_specs = [
        (
            columns["nama"],
            "10 Tanaman dengan Data Terbanyak",
            "Jumlah Data",
            f"{key_prefix}_plant_chart",
        ),
        (
            columns["senyawa"],
            "10 Senyawa Bioaktif yang Paling Sering Muncul",
            "Jumlah Kemunculan",
            f"{key_prefix}_compound_chart",
        ),
        (
            columns["khasiat"],
            "10 Khasiat/Efek yang Paling Sering Muncul",
            "Jumlah Kemunculan",
            f"{key_prefix}_effect_chart",
        ),
    ]

    rendered = 0

    for column, title, x_title, chart_key in chart_specs:
        fig = make_top_value_chart(
            df_data,
            column,
            title,
            x_title,
        )

        if fig is not None:
            st.plotly_chart(
                fig,
                use_container_width=True,
                key=chart_key,
                config={"displaylogo": False},
            )
            rendered += 1

    if rendered == 0:
        st.info(
            "Grafik belum dapat dibuat karena kolom yang diperlukan "
            "belum ditemukan pada dataset."
        )


def render_all_outputs(
    result,
    image_path,
    dataset_text,
    document_text,
    match_text,
    key_prefix,
):
    render_status_box(dataset_text, document_text, match_text)
    render_result_cards(result)
    render_image_section(result, image_path)
    render_relation_table(result, key_prefix=f"{key_prefix}_relation")
    render_kg_section(
        result,
        chart_key=f"{key_prefix}_kg_chart",
    )


def render_preview_kg():
    sample = st.session_state.last_result

    if sample is None:
        sample = {
            "Nama Tanaman": "Serai",
            "Nama Latin": "Cymbopogon nardus",
            "Nama Lokal/Daerah": "Sereh",
            "Bagian Tanaman": "Batang",
            "Zat Bioaktif": "Citronellal, geraniol",
            "Khasiat/Efek Terapeutik": "Antimikroba",
            "Cara Pengolahan": "Digeprek lalu direbus",
            "Komposisi/Dosis": "1–2 batang",
            "Sumber Data": "Buku Saku TOGA Kemenkes RI",
            "Kategori Penyakit": "Belum terdeteksi",
            "Gambar": "Belum terdeteksi",
        }

    render_kg_section(sample, chart_key="dashboard_preview_kg")


def render_quick_access():
    st.markdown(
        '<div class="section-title">⚡ Akses Cepat</div>',
        unsafe_allow_html=True,
    )

    quick_items = [
        (
            "📋",
            "Hasil Ekstraksi Entitas",
            "Lihat kembali entitas hasil ekstraksi.",
            "📋 Hasil Ekstraksi Entitas",
            "quick_entity",
        ),
        (
            "🔗",
            "Bioactive Relation Extraction",
            "Lihat relasi antara tanaman dan atributnya.",
            "🔗 Bioactive Relation Extraction",
            "quick_relation",
        ),
        (
            "🕸️",
            "HerbKG 2.0",
            "Jelajahi Knowledge Graph tanaman herbal.",
            "🕸️ HerbKG 2.0 Explorer",
            "quick_kg",
        ),
        (
            "📦",
            "Downstream",
            "Buka seluruh analisis lanjutan.",
            "📦 Ringkasan Downstream",
            "quick_downstream",
        ),
    ]

    cols = st.columns(4)

    for column, item in zip(cols, quick_items):
        icon, title, description, page_name, key = item

        with column:
            st.markdown(
                f"""<div class="quick-card">
<div style="font-size:2.2rem;">{icon}</div>
<h4>{safe_text(title)}</h4>
<p>{safe_text(description)}</p>
</div>""",
                unsafe_allow_html=True,
            )
            st.button(
                f"Buka {title}",
                use_container_width=True,
                key=key,
                on_click=set_page,
                args=(page_name,),
            )


def render_downstream_overview():
    st.markdown(
        '<div class="section-title">📦 Aplikasi Downstream</div>',
        unsafe_allow_html=True,
    )
    st.write(
        "Aplikasi downstream menggunakan entitas dan relasi hasil ekstraksi "
        "untuk analisis deskriptif, penelusuran bukti, analisis kemiripan, "
        "dan rekomendasi herbal."
    )

    items = [
        (
            "📊 Analisis Deskriptif",
            "Ringkasan distribusi tanaman, senyawa, dan khasiat.",
            "📊 Analisis Deskriptif",
            "down_open_descriptive",
        ),
        (
            "🔎 Query Graf Berbasis Bukti",
            "Pilih tanaman dan telusuri relasi beserta sumber datanya.",
            "🔎 Query Graf Berbasis Bukti",
            "down_open_evidence",
        ),
        (
            "🧬 Analisis Kemiripan",
            "Bandingkan tanaman berdasarkan senyawa dan khasiat.",
            "🧬 Analisis Kemiripan",
            "down_open_similarity",
        ),
        (
            "💡 Rekomendasi Herbal",
            "Cari tanaman berdasarkan kata kunci khasiat atau penyakit.",
            "💡 Rekomendasi Herbal",
            "down_open_recommendation",
        ),
    ]

    cols = st.columns(4)

    for col, item in zip(cols, items):
        title, description, page_name, key = item

        with col:
            st.markdown(
                f"""<div class="quick-card">
<h4>{safe_text(title)}</h4>
<p>{safe_text(description)}</p>
</div>""",
                unsafe_allow_html=True,
            )
            st.button(
                "Buka Fitur",
                use_container_width=True,
                key=key,
                on_click=set_page,
                args=(page_name,),
            )


def render_evidence_query():
    st.markdown(
        '<div class="section-title">🔎 Query Graf Berbasis Bukti</div>',
        unsafe_allow_html=True,
    )

    plants = unique_plant_names(df_data)

    if not plants:
        st.warning("Nama tanaman belum ditemukan pada dataset.")
        return

    selected = st.selectbox(
        "Pilih tanaman",
        plants,
        key="evidence_selected_plant",
    )

    run_query = st.button(
        "🔎 Jalankan Query Graf",
        use_container_width=True,
        key="evidence_run_query",
    )

    if run_query:
        rows = find_rows_by_plant(df_data, selected)
        result = merge_rows_to_result(rows, df_data)

        st.success(
            f"Ditemukan {len(rows):,} baris bukti untuk tanaman {selected}."
        )
        render_result_cards(result)
        render_relation_table(result, key_prefix="evidence_relation")
        render_kg_section(
            result,
            chart_key=f"evidence_kg_{slugify_filename(selected)}",
        )

        st.markdown(
            '<div class="section-title">📚 Data Bukti</div>',
            unsafe_allow_html=True,
        )
        st.dataframe(rows, use_container_width=True, hide_index=True)


def render_similarity_page():
    st.markdown(
        '<div class="section-title">🧬 Analisis Kemiripan Tanaman</div>',
        unsafe_allow_html=True,
    )

    plants = unique_plant_names(df_data)

    if not plants:
        st.warning("Nama tanaman belum ditemukan pada dataset.")
        return

    selected = st.selectbox(
        "Pilih tanaman acuan",
        plants,
        key="similarity_selected_plant",
    )

    top_n = st.slider(
        "Jumlah hasil yang ditampilkan",
        min_value=3,
        max_value=20,
        value=10,
        key="similarity_top_n",
    )

    if st.button(
        "🧬 Hitung Kemiripan",
        use_container_width=True,
        key="similarity_run",
    ):
        similarity_df = calculate_similarity_table(df_data, selected)

        if similarity_df.empty:
            st.info(
                "Kemiripan belum dapat dihitung. Pastikan kolom senyawa "
                "atau khasiat tersedia dan memiliki isi."
            )
            return

        top_df = similarity_df.head(top_n)
        st.dataframe(top_df, use_container_width=True, hide_index=True)

        fig = px.bar(
            top_df.sort_values("Skor Kemiripan"),
            x="Skor Kemiripan",
            y="Tanaman Pembanding",
            orientation="h",
            text="Skor Kemiripan",
            title=f"Tanaman yang Mirip dengan {selected}",
        )
        fig.update_layout(
            height=450,
            plot_bgcolor="white",
            paper_bgcolor="white",
        )
        st.plotly_chart(
            fig,
            use_container_width=True,
            key=f"similarity_chart_{slugify_filename(selected)}",
            config={"displaylogo": False},
        )

        st.download_button(
            "⬇️ Unduh Hasil Kemiripan",
            data=similarity_df.to_csv(index=False).encode("utf-8"),
            file_name=f"kemiripan_{slugify_filename(selected)}.csv",
            mime="text/csv",
            key="similarity_download",
        )


def render_recommendation_page():
    st.markdown(
        '<div class="section-title">💡 Rekomendasi Herbal</div>',
        unsafe_allow_html=True,
    )

    keyword = st.text_input(
        "Masukkan kata kunci khasiat, penyakit, atau senyawa",
        placeholder="Contoh: antimikroba, hipertensi, diabetes, flavonoid",
        key="recommendation_keyword",
    )

    top_n = st.slider(
        "Jumlah rekomendasi",
        min_value=5,
        max_value=30,
        value=10,
        key="recommendation_top_n",
    )

    if st.button(
        "💡 Cari Rekomendasi",
        use_container_width=True,
        key="recommendation_run",
    ):
        recommendations = search_recommendations(df_data, keyword)

        if recommendations.empty:
            st.warning(
                "Belum ditemukan rekomendasi yang cocok dengan kata kunci tersebut."
            )
            return

        top_results = recommendations.head(top_n)
        st.success(
            f"Ditemukan {len(recommendations):,} kandidat rekomendasi."
        )
        st.dataframe(
            top_results,
            use_container_width=True,
            hide_index=True,
        )

        chart_data = (
            top_results.groupby("Nama Tanaman", as_index=False)
            ["Skor Rekomendasi"].max()
            .sort_values("Skor Rekomendasi")
        )

        fig = px.bar(
            chart_data,
            x="Skor Rekomendasi",
            y="Nama Tanaman",
            orientation="h",
            text="Skor Rekomendasi",
            title=f"Rekomendasi Herbal untuk Kata Kunci: {keyword}",
        )
        fig.update_layout(
            height=450,
            plot_bgcolor="white",
            paper_bgcolor="white",
        )
        st.plotly_chart(
            fig,
            use_container_width=True,
            key="recommendation_chart",
            config={"displaylogo": False},
        )

        st.download_button(
            "⬇️ Unduh Rekomendasi CSV",
            data=recommendations.to_csv(index=False).encode("utf-8"),
            file_name="rekomendasi_herbal.csv",
            mime="text/csv",
            key="recommendation_download",
        )


def render_training_page():
    st.markdown(
        '<div class="section-title">🧩 Training Model</div>',
        unsafe_allow_html=True,
    )
    st.info(
        "Halaman ini memeriksa kesiapan dataset dan menyiapkan konfigurasi "
        "training. Training transformer penuh sebaiknya dijalankan di "
        "Google Colab/GPU, bukan langsung di Streamlit CPU."
    )

    model_name = st.selectbox(
        "Pilih model",
        [
            "IndoBERT",
            "BioBERT",
            "SciBERT",
            "BERT Multilingual",
            "RoBERTa",
            "DistilBERT",
        ],
        key="training_model",
    )

    task_name = st.selectbox(
        "Pilih tugas",
        [
            "Named Entity Disambiguation (NED)",
            "Bioactive Information Extraction (BIE)",
            "Relation Extraction (RE)",
            "Klasifikasi Khasiat",
        ],
        key="training_task",
    )

    epochs = st.number_input(
        "Epoch",
        min_value=1,
        max_value=50,
        value=5,
        key="training_epochs",
    )

    batch_size = st.selectbox(
        "Batch size",
        [4, 8, 16, 32],
        index=2,
        key="training_batch",
    )

    learning_rate = st.selectbox(
        "Learning rate",
        [5e-5, 3e-5, 2e-5, 1e-5],
        index=1,
        format_func=lambda value: f"{value:.0e}",
        key="training_lr",
    )

    validate_col, prepare_col = st.columns(2)

    with validate_col:
        validate = st.button(
            "✅ Validasi Kesiapan Dataset",
            use_container_width=True,
            key="training_validate",
        )

    with prepare_col:
        prepare = st.button(
            "⚙️ Siapkan Konfigurasi Training",
            use_container_width=True,
            key="training_prepare",
        )

    if validate:
        required = {
            "Nama Tanaman": columns["nama"],
            "Nama Latin": columns["latin"],
            "Nama Lokal/Daerah": columns["lokal"],
            "Zat Bioaktif": columns["senyawa"],
            "Khasiat/Efek": columns["khasiat"],
        }

        validation_df = pd.DataFrame([
            {
                "Komponen": name,
                "Kolom Dataset": column or "Tidak ditemukan",
                "Status": "Siap" if column else "Perlu diperbaiki",
            }
            for name, column in required.items()
        ])

        st.dataframe(
            validation_df,
            use_container_width=True,
            hide_index=True,
        )

        if all(required.values()):
            st.success("Dataset siap digunakan untuk persiapan training.")
        else:
            st.warning(
                "Beberapa kolom belum ditemukan. Sesuaikan nama kolom pada dataset."
            )

    if prepare:
        config = pd.DataFrame([{
            "Model": model_name,
            "Tugas": task_name,
            "Epoch": epochs,
            "Batch Size": batch_size,
            "Learning Rate": learning_rate,
            "Dataset": st.session_state.dataset_file or "Belum dipilih",
            "Jumlah Baris": len(df_data),
        }])

        st.success("Konfigurasi training berhasil disiapkan.")
        st.dataframe(config, use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ Unduh Konfigurasi Training CSV",
            data=config.to_csv(index=False).encode("utf-8"),
            file_name="konfigurasi_training_hytbionex.csv",
            mime="text/csv",
            key="training_download_config",
        )


def render_settings_page():
    st.markdown(
        '<div class="section-title">⚙️ Pengaturan Sistem</div>',
        unsafe_allow_html=True,
    )

    files = discover_excel_files()

    if not files:
        st.warning(
            "Belum ada file Excel pada repository. Unggah dataset .xlsx "
            "ke folder utama aplikasi."
        )
        return

    current_index = (
        files.index(st.session_state.dataset_file)
        if st.session_state.dataset_file in files
        else 0
    )

    selected_dataset = st.selectbox(
        "Pilih dataset aktif",
        files,
        index=current_index,
        key="settings_dataset_select",
    )

    if st.button(
        "✅ Terapkan Dataset",
        use_container_width=True,
        key="settings_apply_dataset",
    ):
        st.session_state.dataset_file = selected_dataset
        load_dataset.clear()
        st.success(f"Dataset aktif diubah menjadi: {selected_dataset}")
        st.rerun()

    if st.button(
        "🧹 Hapus Cache Dataset",
        use_container_width=True,
        key="settings_clear_cache",
    ):
        load_dataset.clear()
        st.success("Cache dataset berhasil dihapus.")
        st.rerun()

    st.write("**Dataset saat ini:**", st.session_state.dataset_file)
    st.write("**Status:**", dataset_status)
    st.write("**Folder gambar:**", ASSET_DIR)


def render_about_page():
    st.markdown(
        '<div class="section-title">ℹ️ Tentang HyTBIONEX</div>',
        unsafe_allow_html=True,
    )
    st.write(
        """
        **HyTBIONEX** adalah prototipe sistem ekstraksi informasi bioaktif
        tanaman herbal Indonesia. Sistem menghubungkan entitas tanaman,
        nama Latin, nama lokal/daerah, bagian tanaman, zat bioaktif,
        khasiat/efek terapeutik, cara pengolahan, dosis/komposisi, dan
        sumber data ke dalam HerbKG 2.0.

        **Pipeline:** Input → Preprocessing → NED → BIE → Relation Extraction
        → HerbKG 2.0 → Aplikasi Downstream.

        **Peneliti:** Nazwita, M.Kom.
        """
    )


# =========================================================
# ROUTING HALAMAN
# =========================================================
page = st.session_state.page

if page == "🏠 Dashboard":
    left, right = st.columns([2.2, 1])

    with left:
        st.markdown(
            """<div class="hero-banner">
<h2>Selamat Datang di HyTBIONEX</h2>
<p>
Platform cerdas untuk ekstraksi informasi bioaktif tanaman herbal Indonesia,
visualisasi HerbKG 2.0, analisis deskriptif, query graf berbasis bukti,
analisis kemiripan, dan rekomendasi herbal.
</p>
</div>""",
            unsafe_allow_html=True,
        )

    with right:
        st.success("🟢 Sistem siap digunakan")
        st.write("**Dataset aktif**")
        st.caption(dataset_status)
        st.write("**Model pipeline**")
        st.caption("NED → BIE → RE → HerbKG 2.0")

    render_metrics()
    render_analysis_form(
        prefix="dashboard",
        allow_text=True,
        allow_upload=True,
    )
    render_quick_access()
    render_preview_kg()

elif page == "🌿 Input Tanaman":
    render_analysis_form(
        prefix="input_page",
        allow_text=True,
        allow_upload=False,
    )

elif page == "📄 Upload Dokumen":
    render_analysis_form(
        prefix="upload_page",
        allow_text=False,
        allow_upload=True,
    )

elif page == "📋 Hasil Isolasi Entitas":
    if st.session_state.last_result:
        render_status_box(
            st.session_state.last_status.get("dataset", dataset_status),
            st.session_state.last_status.get(
                "document", "Tidak ada dokumen yang diunggah."
            ),
            st.session_state.last_status.get(
                "match", "Belum ada status pencocokan."
            ),
        )
        render_result_cards(st.session_state.last_result)
        render_image_section(
            st.session_state.last_result,
            st.session_state.last_image,
        )
    else:
        st.warning(
            "Belum ada hasil ekstraksi. Jalankan Proses Ekstraksi terlebih dahulu."
        )

elif page == "🔗 Relation Extraction":
    if st.session_state.last_result:
        render_relation_table(
            st.session_state.last_result,
            key_prefix="relation_page",
        )
    else:
        st.warning(
            "Belum ada hasil relasi. Jalankan Proses Ekstraksi terlebih dahulu."
        )

elif page == "🕸️ HerbKG 2.0 Explorer":
    if st.session_state.last_result:
        render_kg_section(
            st.session_state.last_result,
            chart_key="kg_explorer_main",
        )
    else:
        st.info(
            "Belum ada hasil ekstraksi. Berikut contoh Knowledge Graph."
        )
        render_preview_kg()

elif page == "📦 Ringkasan Downstream":
    render_downstream_overview()

elif page == "📊 Analisis Deskriptif":
    render_metrics()
    render_descriptive_charts(key_prefix="descriptive_page")

elif page == "🔎 Query Graf Berbasis Bukti":
    render_evidence_query()

elif page == "🧬 Analisis Kemiripan":
    render_similarity_page()

elif page == "💡 Rekomendasi Herbal":
    render_recommendation_page()

elif page == "📈 Statistik & Analitik":
    render_metrics()
    render_descriptive_charts(key_prefix="statistics_page")
    st.markdown(
        '<div class="section-title">📋 Cuplikan Dataset</div>',
        unsafe_allow_html=True,
    )
    if df_data.empty:
        st.warning("Dataset belum terbaca.")
    else:
        st.dataframe(
            df_data.head(50),
            use_container_width=True,
            hide_index=True,
        )
        st.download_button(
            "⬇️ Unduh Cuplikan Dataset CSV",
            data=df_data.head(500).to_csv(index=False).encode("utf-8"),
            file_name="cuplikan_dataset_hytbionex.csv",
            mime="text/csv",
            key="statistics_download",
        )

elif page == "🧩 Training Model":
    render_training_page()

elif page == "⚙️ Pengaturan":
    render_settings_page()

elif page == "ℹ️ Tentang Aplikasi":
    render_about_page()

else:
    st.error("Halaman tidak ditemukan.")
