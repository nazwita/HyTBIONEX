import os
import re
import html
import time
import pandas as pd
import gradio as gr

DEFAULT_DATASET = "Data set 20098+ Gambar.xlsx"


# =========================================================
# FUNGSI DASAR
# =========================================================

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


# =========================================================
# LOAD DATASET
# =========================================================

def find_dataset_file():
    if os.path.exists(DEFAULT_DATASET):
        return DEFAULT_DATASET

    for f in os.listdir("."):
        lower = f.lower()
        if lower.endswith((".xlsx", ".xls")) and "data" in lower:
            return f

    for f in os.listdir("."):
        if f.lower().endswith((".xlsx", ".xls")):
            return f

    return ""


def load_dataset():
    dataset_path = find_dataset_file()

    if not dataset_path:
        return pd.DataFrame(), "Dataset Excel belum ditemukan"

    try:
        sheets = pd.read_excel(dataset_path, sheet_name=None)
        frames = []

        for _, df in sheets.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                frames.append(df.fillna(""))

        if frames:
            final_df = pd.concat(frames, ignore_index=True).fillna("")
            return final_df, f"Dataset terbaca: {dataset_path} ({len(final_df)} baris)"

        return pd.DataFrame(), f"Dataset kosong: {dataset_path}"

    except Exception as e:
        return pd.DataFrame(), f"Gagal membaca dataset: {e}"


# =========================================================
# BACA DOKUMEN
# =========================================================

def read_pdf_text(path):
    text = ""
    try:
        from pypdf import PdfReader
        reader = PdfReader(path)
        for page in reader.pages:
            text += " " + (page.extract_text() or "")
    except Exception:
        text = ""
    return text


def read_txt_text(path):
    try:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            return f.read()
    except Exception:
        try:
            with open(path, "r", encoding="latin1", errors="ignore") as f:
                return f.read()
        except Exception:
            return ""


def read_csv_text(path):
    try:
        df = pd.read_csv(path).fillna("")
        return " ".join(df.astype(str).values.flatten())
    except Exception:
        try:
            df = pd.read_csv(path, encoding="latin1").fillna("")
            return " ".join(df.astype(str).values.flatten())
        except Exception:
            return ""


def read_excel_text(path):
    text = ""
    try:
        sheets = pd.read_excel(path, sheet_name=None)
        for _, df in sheets.items():
            df = df.fillna("")
            text += " " + " ".join(df.astype(str).values.flatten())
    except Exception:
        text = ""
    return text


def read_file_by_path(path):
    if not path:
        return "", "Tidak ada file dokumen"

    lower = path.lower()

    if lower.endswith(".pdf"):
        text = read_pdf_text(path)
        status = f"PDF terbaca: {os.path.basename(path)} ({len(text)} karakter)"
        if not text.strip():
            status += " | PDF mungkin scan/gambar, sehingga teks tidak terbaca."
        return text, status

    if lower.endswith(".txt"):
        text = read_txt_text(path)
        return text, f"TXT terbaca: {os.path.basename(path)} ({len(text)} karakter)"

    if lower.endswith(".csv"):
        text = read_csv_text(path)
        return text, f"CSV terbaca: {os.path.basename(path)} ({len(text)} karakter)"

    if lower.endswith((".xlsx", ".xls")):
        text = read_excel_text(path)
        return text, f"Excel dokumen terbaca: {os.path.basename(path)} ({len(text)} karakter)"

    return "", "Format dokumen belum didukung"


def find_repository_documents():
    candidates = []
    folders = [".", "dokumen", "documents", "artikel", "data"]

    for folder in folders:
        if not os.path.exists(folder):
            continue

        for name in os.listdir(folder):
            path = os.path.join(folder, name)
            low = path.lower()

            if os.path.isfile(path) and low.endswith((".pdf", ".txt")):
                if "readme" not in low and "requirements" not in low:
                    candidates.append(path)

    return candidates


def read_repository_documents():
    files = find_repository_documents()

    if not files:
        return "", "Tidak ada PDF/TXT otomatis di repository"

    all_text = ""
    used_names = []

    for path in files:
        text, _ = read_file_by_path(path)
        if text.strip():
            all_text += " " + text
            used_names.append(os.path.basename(path))

    if all_text.strip():
        return all_text, f"Dokumen otomatis terbaca dari Files: {', '.join(used_names)} ({len(all_text)} karakter)"

    return "", "PDF/TXT ada, tetapi teks tidak terbaca"


def read_uploaded_or_repository_document(file_path):
    if file_path:
        return read_file_by_path(file_path)

    return read_repository_documents()


# =========================================================
# GAMBAR TANAMAN
# =========================================================

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


# =========================================================
# MATCHING ENTITAS
# =========================================================

def score_match(row, search_text):
    nama_tanaman = clean_text(get_col(row, ["Nama Tanaman", "Nama_Tanaman", "Tanaman", "Nama"]))
    nama_latin = clean_text(get_col(row, ["Nama Latin", "Nama_Latin", "Latin"]))
    nama_lokal = clean_text(get_col(row, ["Nama Lokal/Daerah", "Nama Lokal", "Nama_Daerah", "Bahasa_Daerah"]))

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
        return None, "Dataset kosong atau belum terbaca"

    if not search_text:
        return None, "Input dan dokumen masih kosong"

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

    return None, "Tidak ditemukan kecocokan entitas tanaman pada dataset"


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
        "Nama Lokal/Daerah": get_col(row, ["Nama Lokal/Daerah", "Nama Lokal", "Nama_Daerah", "Bahasa_Daerah"]),
        "Bagian Tanaman": get_col(row, ["Bagian Tanaman", "Bagian_Tanaman", "Bagian Digunakan", "Bagian_Digunakan"]),
        "Zat Bioaktif": get_col(row, ["Zat Bioaktif", "Senyawa Bioaktif", "Senyawa_Bioaktif", "Compound", "Senyawa", "Kandungan"]),
        "Khasiat/Efek Terapeutik": get_col(row, ["Khasiat/Efek Terapeutik", "Khasiat", "Benefit", "Biological_Activity", "Aktivitas Farmakologis"]),
        "Cara Pengolahan": get_col(row, ["Cara Pengolahan", "Cara_Pengolahan", "Pengolahan"]),
        "Komposisi/Dosis": get_col(row, ["Komposisi/Dosis", "Komposisi /Dosis", "Dosis", "Komposisi"]),
        "Sumber Data": get_col(row, ["Sumber Data", "Sumber_Data", "Sumber", "Referensi"])
    }


# =========================================================
# HTML OUTPUT
# =========================================================

def progress_html(percent, message):
    return f"""
    <div class="progress-card">
        <div class="progress-top">
            <span>{safe(message)}</span>
            <b>{percent}%</b>
        </div>
        <div class="progress-track">
            <div class="progress-fill" style="width:{percent}%;"></div>
        </div>
    </div>
    """


def field_card(icon, title, value):
    value_show = safe(value) if value else "Belum terdeteksi"
    return f"""
    <div class="field-card">
        <div class="field-icon">{icon}</div>
        <div class="field-body">
            <div class="field-title">{safe(title)}</div>
            <div class="field-value">{value_show}</div>
        </div>
    </div>
    """


def build_result_html(dataset_status, doc_status, match_status, result):
    status_html = f"""
    <div class="white-card" id="status-section">
        <h2>📌 Status Sistem</h2>
        <p><b>Status Dataset:</b> {safe(dataset_status)}</p>
        <p><b>Status Dokumen:</b> {safe(doc_status)}</p>
        <p><b>Status Koneksi Entitas:</b> {safe(match_status)}</p>
    </div>
    """

    extraction_html = f"""
    <div class="white-card" id="result-section">
        <h2>🌿 Hasil Ekstraksi Informasi Bioaktif</h2>
        <div class="info-grid">
            {field_card("🌿", "Nama Tanaman", result["Nama Tanaman"])}
            {field_card("🔬", "Nama Latin", result["Nama Latin"])}
            {field_card("🇮🇩", "Nama Lokal/Daerah", result["Nama Lokal/Daerah"])}
            {field_card("🍃", "Bagian Tanaman", result["Bagian Tanaman"])}
            {field_card("🧪", "Zat Bioaktif", result["Zat Bioaktif"])}
            {field_card("💚", "Khasiat/Efek Terapeutik", result["Khasiat/Efek Terapeutik"])}
            {field_card("☕", "Cara Pengolahan", result["Cara Pengolahan"])}
            {field_card("⚖️", "Komposisi/Dosis", result["Komposisi/Dosis"])}
            {field_card("📚", "Sumber Data", result["Sumber Data"])}
        </div>
    </div>
    """

    return status_html + extraction_html


def build_relation_html(result):
    return f"""
    <div class="white-card" id="relation-section">
        <h2>🔗 Bioactive Relation Extraction</h2>
        <p><b>{safe(result["Nama Tanaman"])}</b> → <span class="rel">has_latin_name</span> → {safe(result["Nama Latin"])}</p>
        <p><b>{safe(result["Nama Tanaman"])}</b> → <span class="rel">has_local_name</span> → {safe(result["Nama Lokal/Daerah"])}</p>
        <p><b>{safe(result["Nama Tanaman"])}</b> → <span class="rel">uses_part</span> → {safe(result["Bagian Tanaman"])}</p>
        <p><b>{safe(result["Nama Tanaman"])}</b> → <span class="rel">contains_bioactive_compound</span> → {safe(result["Zat Bioaktif"])}</p>
        <p><b>{safe(result["Nama Tanaman"])}</b> → <span class="rel">has_therapeutic_effect</span> → {safe(result["Khasiat/Efek Terapeutik"])}</p>
        <p><b>{safe(result["Nama Tanaman"])}</b> → <span class="rel">processed_by</span> → {safe(result["Cara Pengolahan"])}</p>
        <p><b>{safe(result["Nama Tanaman"])}</b> → <span class="rel">has_dosage</span> → {safe(result["Komposisi/Dosis"])}</p>
        <p><b>{safe(result["Nama Tanaman"])}</b> → <span class="rel">sourced_from</span> → {safe(result["Sumber Data"])}</p>
    </div>
    """


def build_kg_html(result):
    return f"""
    <div class="white-card" id="kg-section">
        <h2>🕸️ Enhanced Herb Knowledge Graph (HerbKG 2.0)</h2>
        <div class="kg-area">
            <svg class="kg-lines" viewBox="0 0 1000 620" preserveAspectRatio="none">
                <line x1="500" y1="310" x2="500" y2="90"></line>
                <line x1="500" y1="310" x2="180" y2="130"></line>
                <line x1="500" y1="310" x2="820" y2="130"></line>
                <line x1="500" y1="310" x2="130" y2="310"></line>
                <line x1="500" y1="310" x2="870" y2="310"></line>
                <line x1="500" y1="310" x2="180" y2="500"></line>
                <line x1="500" y1="310" x2="820" y2="500"></line>
                <line x1="500" y1="310" x2="500" y2="540"></line>
            </svg>

            <div class="kg-node center-node">
                🌿<br>{safe(result["Nama Tanaman"])}<br><small>{safe(result["Nama Latin"])}</small>
            </div>

            <div class="kg-node node-top">🔬<br><b>Nama Latin</b><br>{safe(result["Nama Latin"])}</div>
            <div class="kg-node node-left-top">🇮🇩<br><b>Nama Lokal/Daerah</b><br>{safe(result["Nama Lokal/Daerah"])}</div>
            <div class="kg-node node-right-top">🍃<br><b>Bagian Tanaman</b><br>{safe(result["Bagian Tanaman"])}</div>
            <div class="kg-node node-left-mid">🧪<br><b>Zat Bioaktif</b><br>{safe(result["Zat Bioaktif"])}</div>
            <div class="kg-node node-right-mid">💚<br><b>Khasiat/Efek Terapeutik</b><br>{safe(result["Khasiat/Efek Terapeutik"])}</div>
            <div class="kg-node node-left-bottom">☕<br><b>Cara Pengolahan</b><br>{safe(result["Cara Pengolahan"])}</div>
            <div class="kg-node node-right-bottom">⚖️<br><b>Komposisi/Dosis</b><br>{safe(result["Komposisi/Dosis"])}</div>
            <div class="kg-node node-bottom">📚<br><b>Sumber Data</b><br>{safe(result["Sumber Data"])}</div>
        </div>
    </div>
    """


# =========================================================
# EVENT PROCESS
# =========================================================

def process_extraction(input_text, input_file):
    yield progress_html(5, "Memulai proses ekstraksi..."), "", None, "", ""
    time.sleep(0.2)

    yield progress_html(20, "Membaca dataset herbal..."), "", None, "", ""
    dataset_df, dataset_status = load_dataset()
    time.sleep(0.2)

    yield progress_html(40, "Membaca dokumen PDF/TXT/CSV/Excel..."), "", None, "", ""
    doc_text_raw, doc_status = read_uploaded_or_repository_document(input_file)
    time.sleep(0.2)

    yield progress_html(60, "Mencocokkan entitas tanaman dengan dataset..."), "", None, "", ""
    search_text = (clean_text(input_text) + " " + clean_text(doc_text_raw)).strip()
    best_row, match_status = find_best_row(dataset_df, search_text)
    time.sleep(0.2)

    yield progress_html(80, "Membangun hasil ekstraksi dan relasi..."), "", None, "", ""
    result = make_result(best_row, input_text)
    image_path = find_image(best_row)
    result_html = build_result_html(dataset_status, doc_status, match_status, result)
    relation_html = build_relation_html(result)
    kg_html = build_kg_html(result)
    time.sleep(0.2)

    yield progress_html(100, "Selesai. Hasil ekstraksi berhasil ditampilkan."), result_html, image_path, relation_html, kg_html


def clear_text_after_upload(file_obj):
    if file_obj is not None:
        return ""
    return gr.update()


def clear_input_text():
    return ""


def clear_uploaded_file():
    return None


# =========================================================
# CSS
# =========================================================

css = """
html, body {
    background: #e4f4eb !important;
    color: #111111 !important;
    scroll-behavior: smooth;
}

.gradio-container {
    max-width: 1700px !important;
    background: linear-gradient(180deg, #e4f4eb 0%, #f3efff 100%) !important;
    color: #111111 !important;
    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif !important;
}

/* HEADER */
.hero {
    background:
        linear-gradient(rgba(255,255,255,0.82), rgba(255,255,255,0.82)),
        url('/file=assets/background.png');
    background-size: cover;
    background-position: center;
    border-radius: 0 0 34px 34px;
    padding: 50px 24px;
    text-align: center;
    margin-bottom: 22px;
    box-shadow: 0 12px 28px rgba(0,0,0,0.10);
}

.hero h1 {
    color: #2d7a31 !important;
    font-size: 60px !important;
    font-weight: 900 !important;
    margin: 0 0 10px 0 !important;
}

.hero h2 {
    color: #111111 !important;
    font-size: 28px !important;
    font-weight: 900 !important;
    margin: 0 0 8px 0 !important;
}

.hero p {
    color: #111111 !important;
    font-size: 18px !important;
    font-weight: 700 !important;
    margin: 0 !important;
}

/* SIDEBAR */
.sidebar {
    background: linear-gradient(180deg, #045c35 0%, #0a7a47 55%, #109455 100%) !important;
    border-radius: 26px;
    padding: 24px 18px;
    min-height: 1180px;
    box-shadow: 0 14px 30px rgba(0,0,0,0.18);
    position: sticky;
    top: 12px;
}

.sidebar h1,
.sidebar h2,
.sidebar p,
.sidebar span,
.sidebar div,
.sidebar a {
    color: #ffffff !important;
}

.sidebar h1 {
    font-size: 28px !important;
    font-weight: 900 !important;
    margin-bottom: 18px !important;
}

.sidebar h2 {
    font-size: 22px !important;
    font-weight: 900 !important;
    margin-top: 24px !important;
    margin-bottom: 12px !important;
}

.menu-item {
    display: block;
    text-decoration: none !important;
    background: rgba(255,255,255,0.16);
    color: #ffffff !important;
    padding: 14px 16px;
    border-radius: 14px;
    margin: 10px 0;
    font-weight: 800;
    font-size: 16px;
    border: 1px solid rgba(255,255,255,0.10);
}

.menu-item:hover {
    background: rgba(255,255,255,0.28) !important;
    transform: translateX(4px);
}

/* WELCOME */
.orange-welcome {
    background: linear-gradient(135deg, #f97316 0%, #fb923c 100%) !important;
    border-radius: 24px !important;
    padding: 24px !important;
    border: 2px solid #ea580c !important;
    box-shadow: 0 12px 28px rgba(249,115,22,0.22);
    margin-bottom: 18px;
}

.orange-welcome h2,
.orange-welcome p,
.orange-welcome b {
    color: #111111 !important;
}

.orange-welcome h2 {
    font-size: 24px !important;
    font-weight: 900 !important;
}

.orange-welcome p {
    font-size: 17px !important;
    font-weight: 800 !important;
    line-height: 1.6 !important;
}

/* CARD */
.summary-card, .white-card {
    background: #f7f0ff !important;
    border-radius: 22px;
    padding: 20px;
    border: 2px solid #d8c7ff;
    box-shadow: 0 10px 24px rgba(0,0,0,0.08);
    margin-bottom: 18px;
}

.summary-card h3,
.summary-card p,
.white-card h2,
.white-card p,
.white-card div,
.white-card span {
    color: #111111 !important;
}

.summary-card h3,
.white-card h2 {
    font-weight: 900 !important;
}

/* PANEL INPUT HIJAU */
.panel-green {
    background: linear-gradient(135deg, #0b7a45 0%, #129157 100%) !important;
    border-radius: 22px !important;
    padding: 18px !important;
    border: 2px solid #0a6a3d !important;
    box-shadow: 0 10px 20px rgba(0,0,0,0.12);
    position: relative !important;
}

/* HEADER DALAM PANEL */
.panel-head {
    align-items: center !important;
    margin-bottom: 14px !important;
}

.panel-title {
    display: inline-block;
    background: #1cc05f !important;
    color: #ffffff !important;
    font-weight: 900 !important;
    font-size: 18px !important;
    padding: 12px 18px !important;
    border-radius: 14px !important;
}

.panel-subtitle {
    color: #ffffff !important;
    font-weight: 800 !important;
    margin-bottom: 12px !important;
}

.panel-green *,
.panel-green label {
    color: #ffffff !important;
}

/* INPUT TEXTAREA LILAC */
.panel-green textarea,
.panel-green input {
    background: #f3e8ff !important;
    color: #111111 !important;
    border: 2px solid #d8b4fe !important;
    border-radius: 18px !important;
    font-size: 17px !important;
}

.panel-green textarea::placeholder,
.panel-green input::placeholder {
    color: #5d506d !important;
    opacity: 1 !important;
}

/* UPLOAD FILE AREA LILAC */
.panel-green [data-testid="file-upload"],
.panel-green [data-testid="file-upload"] *,
.panel-green [data-testid="file"],
.panel-green [data-testid="file"] *,
.panel-green .wrap,
.panel-green .container,
.panel-green .file-preview,
.panel-green .file-preview * {
    background: #f3e8ff !important;
    color: #111111 !important;
    border-color: #d8b4fe !important;
}

/* TOMBOL X SUDUT KANAN ATAS */
.x-btn button {
    width: 42px !important;
    height: 42px !important;
    min-width: 42px !important;
    padding: 0 !important;
    border-radius: 50% !important;
    background: #a855f7 !important;
    color: #ffffff !important;
    font-size: 22px !important;
    font-weight: 900 !important;
    border: 2px solid #ffffff !important;
    box-shadow: 0 6px 14px rgba(0,0,0,0.18);
}

.x-btn button:hover {
    background: #7e22ce !important;
}

/* TOMBOL PROSES */
.process-btn button {
    background: linear-gradient(135deg, #f97316 0%, #fb923c 100%) !important;
    color: #ffffff !important;
    border: none !important;
    border-radius: 16px !important;
    font-weight: 900 !important;
    font-size: 20px !important;
    box-shadow: 0 8px 18px rgba(249,115,22,0.28);
    padding: 16px 20px !important;
}

/* LOADING HORIZONTAL */
.progress-card {
    background: #e7fbf3 !important;
    border: 2px solid #5eead4 !important;
    border-radius: 16px !important;
    padding: 14px 16px !important;
    margin: 14px 0 !important;
    box-shadow: 0 8px 18px rgba(20,184,166,0.18);
}

.progress-top {
    display: flex;
    justify-content: space-between;
    align-items: center;
    margin-bottom: 8px;
}

.progress-top span {
    color: #064e3b !important;
    font-weight: 800 !important;
    font-size: 15px !important;
}

.progress-top b {
    color: #0f766e !important;
    font-weight: 900 !important;
    font-size: 18px !important;
}

.progress-track {
    width: 100%;
    height: 13px;
    background: #ccfbf1 !important;
    border-radius: 999px;
    overflow: hidden;
}

.progress-fill {
    height: 100%;
    background: linear-gradient(90deg, #2dd4bf, #14b8a6, #0d9488) !important;
    border-radius: 999px;
    transition: width 0.4s ease;
}

/* HASIL EKSTRAKSI */
.info-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 16px;
}

.field-card {
    background: #eee4ff !important;
    border: 1px solid #d8b4fe;
    border-radius: 18px;
    padding: 16px;
    min-height: 120px;
    display: flex;
    gap: 14px;
    align-items: flex-start;
}

.field-icon {
    font-size: 28px;
    min-width: 42px;
}

.field-title {
    color: #0a6a3d !important;
    font-size: 15px !important;
    font-weight: 900 !important;
    margin-bottom: 6px;
}

.field-value {
    color: #111111 !important;
    font-size: 16px !important;
    font-weight: 700 !important;
    line-height: 1.5 !important;
    word-break: break-word;
}

.rel {
    color: #0b7a45 !important;
    font-weight: 900 !important;
}

/* KNOWLEDGE GRAPH */
.kg-area {
    position: relative;
    height: 620px;
    border-radius: 24px;
    background: radial-gradient(circle, #f3e8ff, #ffffff);
    border: 2px dashed #c084fc;
    overflow: hidden;
}

.kg-lines {
    position: absolute;
    width: 100%;
    height: 100%;
}

.kg-lines line {
    stroke: #118c52;
    stroke-width: 3;
    stroke-dasharray: 8 8;
}

.kg-node {
    position: absolute;
    border-radius: 50%;
    color: #111111 !important;
    text-align: center;
    font-weight: 900 !important;
    display: flex;
    align-items: center;
    justify-content: center;
    flex-direction: column;
    box-shadow: 0 10px 22px rgba(0,0,0,0.18);
    padding: 14px;
    overflow: hidden;
}

.kg-node *,
.kg-node b,
.kg-node small {
    color: #111111 !important;
}

.center-node {
    width: 185px;
    height: 185px;
    left: calc(50% - 92px);
    top: calc(50% - 92px);
    background: #8ee3b0 !important;
    border: 4px solid #0a7a45;
    font-size: 17px;
}

.node-top, .node-left-top, .node-right-top, .node-left-mid,
.node-right-mid, .node-left-bottom, .node-right-bottom, .node-bottom {
    width: 145px;
    height: 145px;
    font-size: 12px;
    line-height: 1.2;
}

.node-top { background: #ddd6fe !important; left: calc(50% - 72px); top: 2%; }
.node-left-top { background: #dff7ea !important; left: 11%; top: 8%; }
.node-right-top { background: #ffe4b8 !important; right: 11%; top: 8%; }
.node-left-mid { background: #eddcff !important; left: 3%; top: calc(50% - 72px); }
.node-right-mid { background: #ffd9df !important; right: 3%; top: calc(50% - 72px); }
.node-left-bottom { background: #dafbe8 !important; left: 11%; bottom: 8%; }
.node-right-bottom { background: #d8f1ff !important; right: 11%; bottom: 8%; }
.node-bottom { background: #e5e7eb !important; left: calc(50% - 72px); bottom: 2%; }

.image-card {
    background: #f3e8ff !important;
    border: 2px dashed #a855f7 !important;
    border-radius: 18px !important;
    padding: 14px !important;
}

@media (max-width: 980px) {
    .info-grid {
        grid-template-columns: 1fr;
    }
    .hero h1 {
        font-size: 42px !important;
    }
}
"""


# =========================================================
# UI APP
# =========================================================

with gr.Blocks(title="HyTBIONEX") as demo:
    gr.HTML(f"<style>{css}</style>")

    gr.HTML("""
    <div class="hero" id="home-section">
        <h1>🌿 HyTBIONEX</h1>
        <h2>Hybrid Transformer for Bioactive Information Extraction</h2>
        <p>Analisis Bioaktif Tanaman Herbal Indonesia & Enhanced Herb Knowledge Graph 2.0</p>
    </div>
    """)

    with gr.Row():
        with gr.Column(scale=1, min_width=260):
            gr.HTML("""
            <div class="sidebar">
                <h1>🌱 DASHBOARD</h1>
                <a class="menu-item" href="#home-section">🏠 Dashboard Utama</a>
                <a class="menu-item" href="#input-section">🌿 Input Data Tanaman</a>
                <a class="menu-item" href="#input-section">📁 Upload Dokumen</a>
                <a class="menu-item" href="#image-section">🖼️ Lampiran Gambar Tanaman</a>
                <a class="menu-item" href="#result-section">📋 Hasil Ekstraksi Entitas</a>
                <a class="menu-item" href="#relation-section">🔗 Bioactive Relation Extraction</a>
                <a class="menu-item" href="#kg-section">🕸️ HerbKG 2.0 Explorer</a>

                <h2>Advanced Downstream Applications</h2>
                <a class="menu-item" href="#downstream-section">📊 Descriptive Analytics</a>
                <a class="menu-item" href="#downstream-section">🔎 Evidence-Based Graph Query</a>
                <a class="menu-item" href="#downstream-section">🧬 Similarity Analysis</a>
                <a class="menu-item" href="#downstream-section">💊 Herbal Recommendation System</a>
            </div>
            """)

        with gr.Column(scale=4):
            with gr.Row():
                with gr.Column(scale=3):
                    gr.HTML("""
                    <div class="orange-welcome">
                        <h2>Selamat Datang di HyTBIONEX</h2>
                        <p>
                        <b>Platform cerdas untuk isolasi informasi bioaktif tanaman herbal Indonesia</b>
                        berbasis <b>Hybrid Transformer</b>, <b>Bioactive Information Extraction</b>,
                        <b>Named Entity Disambiguation</b>, <b>Relation Extraction</b>, dan
                        <b>Enhanced Herb Knowledge Graph</b>.
                        </p>
                    </div>
                    """)
                with gr.Column(scale=1):
                    gr.HTML("""
                    <div class="summary-card">
                        <h3>Ringkasan Data</h3>
                        <p>🌱 Total Tanaman: <b>20.030</b></p>
                        <p>🧪 Senyawa Bioaktif: <b>Dataset</b></p>
                        <p>💚 Khasiat / Efek Terapeutik: <b>Dataset</b></p>
                        <p>🔗 Relasi Triplet: <b>Otomatis</b></p>
                    </div>
                    """)

            gr.HTML('<div id="input-section"></div>')

            with gr.Row():
                with gr.Column(scale=1, elem_classes="panel-green"):
                    with gr.Row(elem_classes="panel-head"):
                        with gr.Column(scale=8):
                            gr.HTML('<div class="panel-title">🌿 1. Input Data Tanaman</div>')
                        with gr.Column(scale=1, min_width=50):
                            clear_text_btn = gr.Button("×", elem_classes="x-btn")

                    gr.HTML('<div class="panel-subtitle">Masukkan nama tanaman atau kalimat artikel herbal</div>')
                    input_text = gr.Textbox(
                        show_label=False,
                        placeholder="Contoh: Kelor, Sirih, Jahe, Kunyit, atau kalimat artikel herbal",
                        lines=5,
                        container=False
                    )

                with gr.Column(scale=1, elem_classes="panel-green"):
                    with gr.Row(elem_classes="panel-head"):
                        with gr.Column(scale=8):
                            gr.HTML('<div class="panel-title">📁 2. Upload Dokumen Artikel / Dataset</div>')
                        with gr.Column(scale=1, min_width=50):
                            clear_file_btn = gr.Button("×", elem_classes="x-btn")

                    gr.HTML('<div class="panel-subtitle">Upload PDF, TXT, CSV, atau Excel</div>')
                    input_file = gr.File(
                        label="",
                        file_types=[".pdf", ".txt", ".csv", ".xlsx", ".xls"],
                        file_count="single",
                        type="filepath"
                    )

            btn = gr.Button("🔍 PROSES EKSTRAKSI", elem_classes="process-btn")

            output_progress = gr.HTML()
            output_result = gr.HTML()

            gr.HTML('<div id="image-section"></div>')
            gr.Markdown("## 🖼️ Lampiran Gambar Tanaman")
            output_image = gr.Image(
                label="Gambar tanaman otomatis dari metadata dataset",
                type="filepath",
                height=300,
                elem_classes="image-card"
            )

            output_relation = gr.HTML()
            output_kg = gr.HTML()

            gr.HTML("""
            <div class="white-card" id="downstream-section">
                <h2>🧬 Model HyTBIONEX Pipeline</h2>
                <p>
                <b>Domain-Adaptive Pretraining & Fine-Tuning</b> →
                <b>Bioactive Information Extraction (BIE)</b> →
                <b>Named Entity Disambiguation (NED)</b> →
                <b>Hybrid Transformer Relation Extraction</b> →
                <b>Bioactive Relation Extraction</b> →
                <b>Enhanced Herb Knowledge Graph (HerbKG 2.0)</b> →
                <b>Advanced Downstream Applications</b>.
                </p>
                <p><b>Catatan:</b> gambar tanaman hanya sebagai lampiran visual berbasis metadata dataset, bukan klasifikasi citra.</p>
            </div>
            """)

    btn.click(
        fn=process_extraction,
        inputs=[input_text, input_file],
        outputs=[output_progress, output_result, output_image, output_relation, output_kg],
        show_progress="hidden"
    )

    input_file.change(
        fn=clear_text_after_upload,
        inputs=input_file,
        outputs=input_text
    )

    clear_text_btn.click(
        fn=clear_input_text,
        inputs=[],
        outputs=input_text
    )

    clear_file_btn.click(
        fn=clear_uploaded_file,
        inputs=[],
        outputs=input_file
    )

demo.queue().launch(
    server_name="0.0.0.0",
    server_port=7860,
    allowed_paths=["."]
)
