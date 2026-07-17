import os
import re
import html
import base64
import unicodedata
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
APP_BUILD = "METRIK-BIOAKTIF-SINKRON-2026-07-17"
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



def _normalise_document_text(text):
    """Meratakan spasi tanpa menghilangkan batas baris penting."""
    text = str(text or "").replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _unique_preserve_order(values):
    seen = set()
    output = []

    for value in values:
        value = re.sub(r"\s+", " ", str(value or "")).strip(" ,;:.|-\n\t")
        key = value.casefold()

        if not value or key in seen:
            continue

        seen.add(key)
        output.append(value)

    return output

def _first_nonempty(value, fallback="Tidak disebutkan dalam dokumen"):
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    return value if value else fallback


def _extract_year(text):
    years = re.findall(r"\b(?:19|20)\d{2}\b", str(text or "")[:8000])
    if not years:
        return ""

    current_reasonable = [year for year in years if 1900 <= int(year) <= 2100]
    return current_reasonable[0] if current_reasonable else ""


def _detect_article_title(text):
    lines = [
        re.sub(r"\s+", " ", line).strip()
        for line in str(text or "").splitlines()
        if re.sub(r"\s+", " ", line).strip()
    ]

    blocked = {
        "abstract", "abstrak", "introduction", "pendahuluan", "keywords",
        "kata kunci", "journal", "jurnal", "volume", "issue", "issn",
        "doi", "received", "accepted", "published", "department",
        "faculty", "university", "universitas", "corresponding author",
    }

    candidates = []

    for index, line in enumerate(lines[:30]):
        lower = line.casefold()

        if any(token in lower for token in blocked):
            continue
        if "@" in line or re.search(r"https?://|www\.", line, re.IGNORECASE):
            continue
        if len(line) < 18 or len(line) > 260:
            continue
        if len(line.split()) < 4:
            continue
        if re.fullmatch(r"[\d\W_]+", line):
            continue

        score = min(len(line), 160)
        score += max(0, 18 - index)

        if line.endswith("."):
            score -= 15
        if line.count(",") >= 4:
            score -= 15

        candidates.append((score, line))

    return max(candidates, default=(0, ""))[1]


def _looks_like_author_line(line):
    line = re.sub(r"\s+", " ", str(line or "")).strip()
    lower = line.casefold()

    blocked = (
        "abstract", "abstrak", "department", "faculty", "university",
        "universitas", "institute", "institut", "laboratory", "journal",
        "jurnal", "volume", "issue", "doi", "received", "accepted",
        "corresponding", "email", "keywords", "kata kunci",
    )

    if not line or any(token in lower for token in blocked):
        return False
    if "@" in line or len(line) > 220 or len(line) < 4:
        return False
    if re.search(r"\b(?:19|20)\d{2}\b", line):
        return False

    cleaned = re.sub(r"[\d*†‡§#]+", "", line)
    name_tokens = re.findall(r"\b[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'’-]+\b", cleaned)
    separators = bool(re.search(r",|\band\b|\bdan\b|&", cleaned, re.IGNORECASE))

    return len(name_tokens) >= 2 and (separators or len(name_tokens) <= 8)


def _detect_article_authors(text, title=""):
    lines = [
        re.sub(r"\s+", " ", line).strip()
        for line in str(text or "").splitlines()
        if re.sub(r"\s+", " ", line).strip()
    ]

    for line in lines[:50]:
        match = re.search(
            r"^(?:authors?|penulis)\s*[:\-]\s*(.+)$",
            line,
            flags=re.IGNORECASE,
        )
        if match:
            return re.sub(r"[\d*†‡§#]+", "", match.group(1)).strip(" ,;")

    title_index = -1
    if title:
        title_key = clean_text(title)
        for index, line in enumerate(lines[:30]):
            if clean_text(line) == title_key:
                title_index = index
                break

    search_start = title_index + 1 if title_index >= 0 else 0

    for line in lines[search_start:search_start + 12]:
        if _looks_like_author_line(line):
            return re.sub(r"[\d*†‡§#]+", "", line).strip(" ,;")

    for index, line in enumerate(lines[:60]):
        if "@" in line and index > 0:
            for previous in reversed(lines[max(0, index - 3):index]):
                if _looks_like_author_line(previous):
                    return re.sub(r"[\d*†‡§#]+", "", previous).strip(" ,;")

    return ""


def read_uploaded_document(uploaded_file):
    """
    Membaca artikel secara langsung dan mengambil metadata bibliografi.
    Fungsi ini tidak menggunakan dataset Excel.
    """
    if uploaded_file is None:
        return {}, "Tidak ada dokumen yang diunggah."

    filename = uploaded_file.name
    filename_lower = filename.lower()
    document = {
        "filename": filename,
        "text": "",
        "title": "",
        "authors": "",
        "year": "",
        "pages": 0,
        "format": Path(filename).suffix.lower().replace(".", "").upper(),
    }

    try:
        uploaded_file.seek(0)

        if filename_lower.endswith(".pdf"):
            from pypdf import PdfReader

            reader = PdfReader(uploaded_file)
            page_texts = [(page.extract_text() or "") for page in reader.pages]
            document["text"] = _normalise_document_text("\n\n".join(page_texts))
            document["pages"] = len(reader.pages)

            metadata = reader.metadata or {}
            document["title"] = str(getattr(metadata, "title", "") or "").strip()
            document["authors"] = str(getattr(metadata, "author", "") or "").strip()

            creation_date = getattr(metadata, "creation_date", None)
            if creation_date is not None:
                document["year"] = str(getattr(creation_date, "year", "") or "")

        elif filename_lower.endswith(".docx"):
            from docx import Document

            word_document = Document(uploaded_file)
            document["text"] = _normalise_document_text(
                "\n".join(paragraph.text for paragraph in word_document.paragraphs)
            )
            properties = word_document.core_properties
            document["title"] = str(properties.title or "").strip()
            document["authors"] = str(properties.author or "").strip()
            if properties.created:
                document["year"] = str(properties.created.year)

        elif filename_lower.endswith(".txt"):
            raw = uploaded_file.read()
            document["text"] = _normalise_document_text(
                raw.decode("utf-8", errors="ignore")
            )

        elif filename_lower.endswith(".csv"):
            frame = pd.read_csv(uploaded_file).fillna("")
            document["text"] = _normalise_document_text(
                "\n".join(" | ".join(map(str, row)) for row in frame.values)
            )

        elif filename_lower.endswith((".xlsx", ".xls")):
            sheets = pd.read_excel(uploaded_file, sheet_name=None)
            text_parts = []
            for sheet_name, frame in sheets.items():
                frame = frame.fillna("")
                text_parts.append(f"SHEET: {sheet_name}")
                text_parts.extend(" | ".join(map(str, row)) for row in frame.values)
            document["text"] = _normalise_document_text("\n".join(text_parts))

        else:
            return {}, "Format dokumen belum didukung."

        if not document["text"]:
            return (
                document,
                f"Dokumen {filename} terbaca, tetapi teks tidak ditemukan. "
                "PDF kemungkinan berupa scan/gambar dan membutuhkan OCR.",
            )

        detected_title = _detect_article_title(document["text"])
        if _metadata_value_is_weak(
            document.get("title"),
            filename=filename,
            value_type="title",
        ):
            document["title"] = detected_title

        detected_authors = _detect_article_authors(
            document["text"],
            document.get("title", ""),
        )
        if _metadata_value_is_weak(
            document.get("authors"),
            filename=filename,
            value_type="authors",
        ):
            document["authors"] = detected_authors

        if not document["year"]:
            document["year"] = _extract_year(document["text"])

        status_parts = [
            f"{document['format']} terbaca: {filename}",
            f"{len(document['text']):,} karakter",
        ]
        if document["pages"]:
            status_parts.append(f"{document['pages']} halaman")

        return document, " | ".join(status_parts)

    except Exception as error:
        return {}, f"Gagal membaca dokumen: {error}"


def _sentences(text):
    text = _normalise_document_text(text)
    parts = re.split(r"(?<=[.!?])\s+|\n+", text)
    return [re.sub(r"\s+", " ", part).strip() for part in parts if part.strip()]


def _capture_label_value(text, labels, max_length=220):
    label_pattern = "|".join(re.escape(label) for label in labels)
    patterns = [
        rf"(?:{label_pattern})\s*[:\-]\s*([^\n.;]{{2,{max_length}}})",
        rf"(?:{label_pattern})\s+(?:adalah|yaitu|berupa|meliputi|include(?:s)?|are|is)\s+([^\n.;]{{2,{max_length}}})",
    ]

    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return match.group(1).strip(" ,;:")

    return ""


def _extract_latin_name(text):
    labelled = _capture_label_value(
        text,
        ["nama latin", "nama ilmiah", "scientific name", "botanical name"],
        max_length=100,
    )
    if labelled:
        match = re.search(
            r"\b([A-Z][a-z]{2,}\s+[a-z][a-z-]{2,}(?:\s+(?:subsp\.|var\.)\s+[a-z-]+)?)\b",
            labelled,
        )
        if match:
            return match.group(1)

    pattern = re.compile(
        r"\b([A-Z][a-z]{2,}\s+[a-z][a-z-]{2,}(?:\s+(?:subsp\.|var\.)\s+[a-z-]+)?)\b"
    )
    blocked_first = {
        "The", "This", "Table", "Figure", "Results", "Discussion",
        "Introduction", "Abstract", "Indonesia", "Journal", "Article",
        "Keywords", "Department", "University", "Faculty", "Volume",
        "Received", "Accepted", "Background", "Methods", "Conclusion",
        "Research", "Data", "Study", "Based", "According", "Medicinal",
    }

    candidates = [
        candidate
        for candidate in pattern.findall(text)
        if candidate.split()[0] not in blocked_first
    ]

    if not candidates:
        return ""

    counts = Counter(candidate.casefold() for candidate in candidates)
    display = {}
    for candidate in candidates:
        display.setdefault(candidate.casefold(), candidate)

    best_key = counts.most_common(1)[0][0]
    return display[best_key]


def _clean_plant_name_candidate(value):
    """Membersihkan kandidat nama tanaman agar tidak menjadi kalimat panjang."""
    value = re.sub(r"\s+", " ", str(value or "")).strip(" ,;:.()[]{}|-")
    value = re.sub(
        r"^(?:the|a|an|tanaman|plant|herbal plant|medicinal plant)\s+",
        "",
        value,
        flags=re.IGNORECASE,
    )
    value = re.split(
        r"\b(?:contains?|mengandung|showed|menunjukkan|has|memiliki|"
        r"was|were|is|are|used|digunakan|with|dengan)\b",
        value,
        maxsplit=1,
        flags=re.IGNORECASE,
    )[0].strip(" ,;:.-")

    blocked = {
        "abstract", "abstrak", "introduction", "pendahuluan", "result",
        "results", "discussion", "conclusion", "method", "methods",
        "article", "journal", "research", "study", "penelitian",
        "radikal bebas", "free radical", "antioksidan", "antioxidant",
        "aktivitas antioksidan", "antioxidant activity",
        "senyawa bioaktif", "bioactive compound", "bioactive compounds",
        "ekstrak", "extract", "ekstrak daun", "leaf extract",
    }

    if not value or value.casefold() in blocked:
        return ""
    if len(value) > 80 or len(value.split()) > 7:
        return ""
    if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", value):
        return ""

    return value


def _normalise_lookup_key(value):
    """Normalisasi nama tanaman/nama file agar pencarian gambar lebih toleran."""
    value = unicodedata.normalize("NFKD", str(value or ""))
    value = "".join(character for character in value if not unicodedata.combining(character))
    value = value.casefold()
    value = re.sub(r"[^a-z0-9]+", " ", value)
    return re.sub(r"\s+", " ", value).strip()


def _available_plant_images():
    """Mengambil semua gambar tanaman dari assets, gambar, dan images."""
    image_files = []
    valid_extensions = {".jpg", ".jpeg", ".png", ".webp"}

    for directory in [ASSET_DIR, "gambar", "images"]:
        folder = Path(directory)
        if not folder.exists() or not folder.is_dir():
            continue

        for path in folder.rglob("*"):
            if path.is_file() and path.suffix.casefold() in valid_extensions:
                image_files.append(path)

    return image_files


def _plant_name_from_title_and_assets(article_title):
    """
    Mengambil nama umum tanaman dari judul dengan mencocokkannya terhadap
    nama file gambar. Contoh: judul mengandung 'Kelor' dan tersedia
    assets/kelor.jpg, maka nama tanaman menjadi 'Kelor'.
    """
    title_key = _normalise_lookup_key(article_title)
    if not title_key:
        return ""

    title_tokens = set(title_key.split())
    best_name = ""
    best_score = 0

    for image_path in _available_plant_images():
        stem = re.sub(r"[_-]+", " ", image_path.stem).strip()
        stem_key = _normalise_lookup_key(stem)
        if not stem_key:
            continue

        stem_tokens = set(stem_key.split())
        score = 0

        if re.search(rf"(?<![a-z0-9]){re.escape(stem_key)}(?![a-z0-9])", title_key):
            score = 120 + len(stem_key)
        elif stem_tokens and stem_tokens.issubset(title_tokens):
            score = 100 + len(stem_tokens) * 5
        else:
            overlap = len(stem_tokens & title_tokens)
            if overlap and overlap / max(len(stem_tokens), 1) >= 0.75:
                score = 60 + overlap * 5

        if score > best_score:
            best_score = score
            best_name = stem

    if best_name:
        return " ".join(word.capitalize() for word in best_name.split())

    # Fallback nama tanaman umum yang sering muncul dalam judul artikel.
    common_names = [
        "kelor", "kopi arabika", "kopi robusta", "kopi", "jahe merah", "jahe",
        "kunyit", "temulawak", "serai", "sereh", "sambiloto", "kayu manis",
        "daun salam", "sirih merah", "sirih", "pegagan", "meniran", "mengkudu",
        "jambu biji", "lidah buaya", "kumis kucing", "seledri", "kemangi",
        "kencur", "lengkuas", "bawang putih", "bawang merah", "cengkeh",
        "pala", "kapulaga", "rosella", "mahkota dewa", "binahong", "brotowali",
        "belimbing wuluh", "alpukat", "sirsak", "pepaya", "manggis", "jeruk nipis",
    ]

    for name in sorted(common_names, key=len, reverse=True):
        name_key = _normalise_lookup_key(name)
        if re.search(rf"(?<![a-z0-9]){re.escape(name_key)}(?![a-z0-9])", title_key):
            return " ".join(word.capitalize() for word in name.split())

    return ""


def _extract_plant_name_from_title(article_title, latin_name=""):
    """
    Mengambil nama tanaman dari judul artikel sebagai fallback.
    Digunakan hanya jika nama umum tidak ditemukan pada isi artikel.
    """
    title = re.sub(r"\s+", " ", str(article_title or "")).strip()
    if not title:
        return ""

    if latin_name:
        escaped_latin = re.escape(latin_name)

        patterns = [
            rf"([A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{{2,60}})"
            rf"\s*\(\s*{escaped_latin}\s*\)",
            rf"{escaped_latin}\s*\(\s*"
            rf"([A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{{2,60}})\s*\)",
            rf"([A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{{2,60}}),?\s+"
            rf"(?:scientifically known as|botanically known as)\s+{escaped_latin}",
        ]

        for pattern in patterns:
            match = re.search(pattern, title, flags=re.IGNORECASE)
            if match:
                candidate = _clean_plant_name_candidate(match.group(1))
                if candidate:
                    return candidate

    # Pola judul umum: "... of Kelor", "... pada Jahe", "... from Turmeric"
    title_without_subtitle = re.split(r"\s*[:|]\s*", title, maxsplit=1)[0]
    patterns = [
        r"(?:bioactive compounds?|phytochemical(?:s| screening)?|"
        r"chemical constituents?|antioxidant activity|antimicrobial activity|"
        r"therapeutic potential|medicinal properties|ekstraksi senyawa bioaktif|"
        r"kandungan bioaktif|aktivitas antioksidan|aktivitas antimikroba)"
        r".*?\b(?:of|from|in|pada|dari)\s+"
        r"([A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{2,70})$",
        r"\b(?:of|from|in|pada|dari)\s+"
        r"([A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{2,70})$",
    ]

    for pattern in patterns:
        match = re.search(pattern, title_without_subtitle, flags=re.IGNORECASE)
        if match:
            candidate = _clean_plant_name_candidate(match.group(1))
            if candidate:
                return candidate

    # Judul seperti "Analisis Senyawa Bioaktif Daun Kelor" atau
    # "Aktivitas Antioksidan Kopi" ditangani melalui nama gambar/daftar umum.
    title_plant = _plant_name_from_title_and_assets(title)
    if title_plant:
        return title_plant

    return ""


def _metadata_value_is_weak(value, filename="", value_type="title"):
    """Mendeteksi metadata PDF/DOCX yang generik atau tidak informatif."""
    value = re.sub(r"\s+", " ", str(value or "")).strip()
    if not value:
        return True

    lower = value.casefold()
    filename_stem = Path(str(filename or "")).stem.casefold()

    generic_values = {
        "untitled", "unknown", "anonymous", "admin", "administrator",
        "user", "microsoft office user", "default", "article", "document",
    }

    if lower in generic_values:
        return True
    if filename_stem and lower == filename_stem:
        return True
    if value_type == "title" and len(value.split()) < 3:
        return True
    if value_type == "authors" and len(value) < 4:
        return True

    return False


def _is_missing_value(value):
    value = clean_text(value)
    return (
        not value
        or value in {
            "belum terdeteksi",
            "tidak terdeteksi",
            "tidak disebutkan dalam dokumen",
            "none",
            "nan",
        }
    )





def _clean_sentence_value(value):
    """Membersihkan nilai agar dapat dimasukkan ke dalam satu kalimat."""
    value = str(value or "").strip()
    value = re.sub(r"^\s*(?:1\)|2\)|3\)|[-•])\s*", "", value)
    value = re.sub(r"\s+", " ", value)
    return value.strip(" ,;:.-")


def _capitalise_first(value):
    value = _clean_sentence_value(value)
    if not value:
        return value
    return value[0].upper() + value[1:]


def build_therapeutic_conclusion(result):
    """
    Membuat Keterangan secara langsung, tanpa frasa
    'memiliki efek terapeutik' dan tanpa poin ringkasan artikel.

    Contoh:
    Jahe digunakan sebagai Antipiretik untuk membantu mengatasi demam.
    """
    plant = _clean_sentence_value(
        result.get("Nama Tanaman", "")
    )
    activity = _clean_sentence_value(
        result.get("Khasiat/Efek Terapeutik", "")
    )
    disease = _clean_sentence_value(
        result.get("Kategori Penyakit", "")
    )

    if _is_missing_value(plant):
        plant = "Tanaman tersebut"

    activity_missing = _is_missing_value(activity)
    disease_missing = _is_missing_value(disease)

    if not activity_missing and not disease_missing:
        return (
            f"{plant} digunakan sebagai {_capitalise_first(activity)} "
            f"untuk membantu mengatasi {disease.lower()}."
        )

    if not activity_missing:
        return (
            f"{plant} digunakan sebagai {_capitalise_first(activity)}."
        )

    if not disease_missing:
        return (
            f"{plant} digunakan untuk membantu mengatasi "
            f"{disease.lower()}."
        )

    return (
        f"Khasiat dan kategori penyakit yang berkaitan dengan "
        f"{plant} belum terdeteksi."
    )


LATIN_TO_COMMON_PLANT = {
    "moringa oleifera": "Kelor",
    "syzygium aromaticum": "Cengkeh",
    "zingiber officinale": "Jahe",
    "curcuma longa": "Kunyit",
    "cymbopogon nardus": "Serai",
    "cymbopogon citratus": "Serai",
    "cinnamomum burmannii": "Kayu Manis",
    "allium sativum": "Bawang Putih",
    "piper betle": "Sirih",
    "coffea arabica": "Kopi Arabika",
    "coffea canephora": "Kopi Robusta",
}


def _common_name_from_latin(latin_name):
    latin_key = clean_text(latin_name)
    return LATIN_TO_COMMON_PLANT.get(latin_key, "")

def _extract_plant_name_from_document_context(document, latin_name=""):
    """
    Mengambil nama tanaman terutama dari judul artikel, bagian awal artikel,
    dan nama file. Contoh judul yang memuat Kelor atau Kopi akan langsung
    menghasilkan Nama Tanaman yang sesuai.
    """
    text = str(document.get("text", "") or "")
    title = str(document.get("title", "") or "")
    filename_stem = Path(str(document.get("filename", "") or "")).stem

    header_lines = [
        re.sub(r"\s+", " ", line).strip()
        for line in text.splitlines()[:45]
        if re.sub(r"\s+", " ", line).strip()
    ]

    candidates = [title]
    candidates.extend(header_lines[:20])
    candidates.append(" ".join(header_lines[:12]))
    candidates.append(filename_stem)

    # Prioritas pertama: nama umum yang jelas pada judul atau file gambar.
    for candidate in candidates:
        if not candidate:
            continue

        name = _plant_name_from_title_and_assets(candidate)
        if name:
            return name

    # Prioritas kedua: pemetaan nama Latin ke nama tanaman umum.
    mapped_name = _common_name_from_latin(latin_name)
    if mapped_name:
        return mapped_name

    # Prioritas ketiga: pola bahasa dalam judul.
    for candidate in candidates:
        if not candidate:
            continue

        name = _extract_plant_name_from_title(candidate, latin_name)
        if name:
            return name

    return ""

def _extract_plant_and_local_names(text, latin_name, article_title=""):
    """
    Mengekstrak nama tanaman dan nama lokal dari isi artikel.
    Jika nama umum tidak ditemukan, nama Latin digunakan sebagai fallback
    agar entitas Nama Tanaman tetap tampil pada hasil upload dokumen.
    """
    plant_name = _capture_label_value(
        text,
        [
            "nama tanaman", "nama umum", "common name", "plant name",
            "medicinal plant", "herbal name",
        ],
        max_length=90,
    )
    local_name = _capture_label_value(
        text,
        [
            "nama lokal/daerah", "nama lokal", "nama daerah", "local name",
            "vernacular name", "locally known as", "commonly known as",
            "known as", "dikenal sebagai", "disebut sebagai", "disebut",
        ],
        max_length=100,
    )

    # Pola eksplisit: "Kelor (Moringa oleifera)" atau sebaliknya.
    if latin_name:
        escaped_latin = re.escape(latin_name)

        if not plant_name:
            patterns = [
                rf"\b([A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{{2,60}})"
                rf"\s*\(\s*{escaped_latin}\s*\)",
                rf"{escaped_latin}\s*\(\s*"
                rf"([A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{{2,60}})\s*\)",
                rf"\b([A-ZÀ-ÖØ-öø-ÿ][A-Za-zÀ-ÖØ-öø-ÿ'’ -]{{2,60}}),?\s+"
                rf"(?:scientifically known as|botanically known as)\s+{escaped_latin}",
            ]

            for pattern in patterns:
                match = re.search(pattern, text, flags=re.IGNORECASE)
                if match:
                    plant_name = match.group(1)
                    break

        if not local_name:
            local_patterns = [
                rf"{escaped_latin}.{{0,160}}?"
                rf"(?:locally known as|commonly known as|known as|local name|"
                rf"nama lokal|nama daerah|dikenal sebagai|disebut sebagai|disebut)"
                rf"\s*[:\-]?\s*([A-Za-zÀ-ÖØ-öø-ÿ'’ -]{{2,60}})",
                rf"(?:locally known as|commonly known as|known as|"
                rf"nama lokal|nama daerah|dikenal sebagai|disebut sebagai)"
                rf"\s*[:\-]?\s*([A-Za-zÀ-ÖØ-öø-ÿ'’ -]{{2,60}})"
                rf".{{0,120}}?{escaped_latin}",
            ]

            for pattern in local_patterns:
                match = re.search(
                    pattern,
                    text,
                    flags=re.IGNORECASE | re.DOTALL,
                )
                if match:
                    local_name = match.group(1)
                    break

    plant_name = _clean_plant_name_candidate(plant_name)
    local_name = _clean_plant_name_candidate(local_name)

    # Fallback dari judul artikel.
    if not plant_name:
        plant_name = _extract_plant_name_from_title(article_title, latin_name)

    # Nama Latin tetap ditampilkan sebagai Nama Tanaman bila nama umum
    # memang tidak dituliskan oleh artikel.
    if not plant_name and latin_name:
        plant_name = latin_name

    if local_name:
        local_name = re.split(r"[,;|]", local_name)[0].strip()

    # Nama lokal hanya disamakan dengan nama umum, bukan dengan nama Latin.
    if not local_name and plant_name and clean_text(plant_name) != clean_text(latin_name):
        local_name = plant_name

    return plant_name, local_name


def _extract_plant_parts(text):
    labelled = _capture_label_value(
        text,
        [
            "bagian tanaman", "bagian yang digunakan", "bagian digunakan",
            "plant part", "part used", "parts used",
        ],
        max_length=140,
    )

    part_map = {
        "daun": ["daun", "leaf", "leaves"],
        "akar": ["akar", "root", "roots"],
        "batang": ["batang", "stem", "stems"],
        "kulit batang": ["kulit batang", "bark"],
        "rimpang": ["rimpang", "rhizome", "rhizomes"],
        "umbi": ["umbi", "tuber", "tubers", "bulb", "bulbs"],
        "buah": ["buah", "fruit", "fruits"],
        "biji": ["biji", "seed", "seeds"],
        "bunga": ["bunga", "flower", "flowers"],
        "herba": ["herba", "whole plant", "aerial parts"],
        "getah": ["getah", "latex", "sap"],
    }

    search_text = f"{labelled} {text}" if labelled else text
    found = []

    for canonical, variants in part_map.items():
        for variant in variants:
            if re.search(rf"\b{re.escape(variant)}\b", search_text, re.IGNORECASE):
                found.append(canonical.title())
                break

    return ", ".join(_unique_preserve_order(found)[:6])


def _clean_compound_candidate(candidate):
    candidate = re.sub(
        r"^(?:the|a|an|its|their|senyawa|senyawa bioaktif|bioactive compounds?|phytochemicals?|kandungan kimia)\s+",
        "",
        str(candidate or ""),
        flags=re.IGNORECASE,
    )
    candidate = re.sub(r"\([^)]*\)", "", candidate)
    candidate = re.sub(r"\s+", " ", candidate).strip(" ,;:.|-\n\t")

    blocked = (
        "activity", "aktivitas", "effect", "efek", "extract", "ekstrak",
        "sample", "sampel", "method", "metode", "result", "hasil",
        "plant", "tanaman", "leaf", "leaves", "daun",
    )

    if not candidate or len(candidate) > 80 or len(candidate.split()) > 7:
        return ""
    if any(word in candidate.casefold() for word in blocked):
        return ""
    if not re.search(r"[A-Za-zÀ-ÖØ-öø-ÿ]", candidate):
        return ""

    return candidate



def _extract_bioactive_compounds(text):
    """
    Mengambil hanya senyawa bioaktif yang paling relevan dari artikel.
    Hasil dibatasi maksimal delapan senyawa agar tidak menjadi uraian panjang.
    """
    text = str(text or "")
    compounds = []

    known_compounds = [
        "quercetin", "kaempferol", "curcumin", "gingerol", "shogaol",
        "eugenol", "citronellal", "citronellol", "geraniol", "limonene",
        "linalool", "menthol", "thymol", "carvacrol", "cinnamaldehyde",
        "cinnamic acid", "gallic acid", "caffeic acid", "chlorogenic acid",
        "ferulic acid", "ellagic acid", "catechin", "epicatechin",
        "rutin", "apigenin", "luteolin", "naringenin", "hesperidin",
        "anthocyanin", "antosianin", "beta carotene", "β-carotene",
        "ascorbic acid", "vitamin c", "oleanolic acid", "ursolic acid",
        "andrographolide", "asiaticoside", "madecassoside", "allicin",
        "piperine", "capsaicin", "mangiferin", "xanthone", "acetogenin",
        "flavonoid", "alkaloid", "saponin", "tannin", "tanin",
        "terpenoid", "steroid", "phenolic", "fenolik", "polyphenol",
        "polifenol", "glycoside", "glikosida", "essential oil",
        "minyak atsiri",
    ]

    cue_terms = [
        "senyawa bioaktif", "senyawa aktif", "kandungan kimia",
        "mengandung", "bioactive compound", "active compound",
        "chemical constituent", "phytochemical", "contains",
    ]

    # Prioritaskan senyawa yang muncul pada kalimat yang jelas membahas
    # kandungan bioaktif.
    relevant_sentences = []
    for sentence in _sentences(text):
        lower = sentence.casefold()
        if any(cue in lower for cue in cue_terms):
            relevant_sentences.append(sentence)
        if len(relevant_sentences) >= 12:
            break

    relevant_text = " ".join(relevant_sentences)

    for compound in known_compounds:
        if relevant_text and re.search(
            rf"(?<!\w){re.escape(compound)}(?!\w)",
            relevant_text,
            re.IGNORECASE,
        ):
            compounds.append(compound)

    # Ambil daftar setelah kata pemicu jika nama senyawanya belum tercakup
    # dalam kamus.
    cue_patterns = [
        r"(?:senyawa bioaktif|senyawa aktif|kandungan kimia|mengandung)"
        r"\s*(?:adalah|yaitu|berupa|meliputi|:)?\s*([^.;\n]{3,180})",
        r"(?:bioactive compounds?|active compounds?|chemical constituents?|"
        r"phytochemicals?|contains?)"
        r"\s*(?:include(?:s)?|are|is|such as|:)?\s*([^.;\n]{3,180})",
    ]

    for pattern in cue_patterns:
        for match in re.finditer(pattern, text, flags=re.IGNORECASE):
            segment = re.split(
                r"\b(?:which|that|yang|with|dengan|showed|menunjukkan|"
                r"possess|memiliki|activity|aktivitas)\b",
                match.group(1),
                maxsplit=1,
                flags=re.IGNORECASE,
            )[0]

            for item in re.split(
                r",|;|\band\b|\bdan\b",
                segment,
                flags=re.IGNORECASE,
            ):
                cleaned = _clean_compound_candidate(item)
                if cleaned:
                    compounds.append(cleaned)

            if len(compounds) >= 8:
                break

        if len(compounds) >= 8:
            break

    # Fallback bila kalimat pemicu tidak ditemukan.
    if not compounds:
        compact_text = text[:16000]
        for compound in known_compounds:
            if re.search(
                rf"(?<!\w){re.escape(compound)}(?!\w)",
                compact_text,
                re.IGNORECASE,
            ):
                compounds.append(compound)

    return ", ".join(_unique_preserve_order(compounds)[:8])

def _extract_biological_activities(text):
    activity_map = {
        "Antioksidan": ["antioksidan", "antioxidant"],
        "Antimikroba": ["antimikroba", "antimicrobial"],
        "Antibakteri": ["antibakteri", "antibacterial"],
        "Antijamur": ["antijamur", "antifungal"],
        "Antiinflamasi": ["antiinflamasi", "anti-inflammatory", "anti inflammatory"],
        "Antidiabetes": ["antidiabetes", "antidiabetic", "hypoglycemic"],
        "Antihipertensi": ["antihipertensi", "antihypertensive"],
        "Analgesik": ["analgesik", "analgesic"],
        "Antipiretik": ["antipiretik", "antipyretic"],
        "Antikanker": ["antikanker", "anticancer", "cytotoxic"],
        "Hepatoprotektif": ["hepatoprotektif", "hepatoprotective"],
        "Imunomodulator": ["imunomodulator", "immunomodulatory"],
        "Diuretik": ["diuretik", "diuretic"],
        "Antivirus": ["antivirus", "antiviral"],
        "Gastroprotektif": ["gastroprotektif", "gastroprotective"],
        "Penyembuhan Luka": ["penyembuhan luka", "wound healing"],
    }

    found = []
    for canonical, variants in activity_map.items():
        if any(
            re.search(rf"(?<!\w){re.escape(variant)}(?!\w)", text, re.IGNORECASE)
            for variant in variants
        ):
            found.append(canonical)

    labelled = _capture_label_value(
        text,
        [
            "aktivitas biologis", "efek terapeutik", "khasiat", "manfaat",
            "biological activity", "therapeutic effect", "pharmacological activity",
        ],
        max_length=220,
    )
    if labelled:
        for item in re.split(r",|;|\band\b|\bdan\b", labelled, flags=re.IGNORECASE):
            item = re.sub(r"\s+", " ", item).strip(" ,;:.")
            if 2 <= len(item) <= 70:
                found.append(item)

    return ", ".join(_unique_preserve_order(found)[:12])



def _extract_disease_category(text):
    """
    Mengambil kategori penyakit atau kondisi kesehatan yang disebutkan
    secara langsung dalam dokumen. Hasil dibuat singkat dan unik.
    """
    labelled = _capture_label_value(
        text,
        [
            "kategori penyakit",
            "nama penyakit",
            "penyakit",
            "disease category",
            "disease",
            "health condition",
            "medical condition",
        ],
        max_length=180,
    )

    disease_map = {
        "Gangguan Pencernaan": [
            "sakit perut", "nyeri perut", "gangguan pencernaan",
            "gastritis", "maag", "diare", "disentri", "kembung",
            "mual", "ulkus lambung", "peptic ulcer",
        ],
        "Diabetes": [
            "diabetes", "diabetes mellitus", "hiperglikemia",
            "hyperglycemia", "gula darah",
        ],
        "Hipertensi": [
            "hipertensi", "hypertension", "darah tinggi",
        ],
        "Peradangan": [
            "inflamasi", "inflammation", "radang",
        ],
        "Infeksi Bakteri": [
            "infeksi bakteri", "bacterial infection",
        ],
        "Infeksi Jamur": [
            "infeksi jamur", "fungal infection",
        ],
        "Kanker": [
            "kanker", "cancer", "tumor",
        ],
        "Demam": [
            "demam", "fever",
        ],
        "Nyeri": [
            "nyeri", "pain", "sakit kepala", "headache",
        ],
        "Penyakit Hati": [
            "penyakit hati", "liver disease", "hepatitis",
        ],
        "Penyakit Kulit": [
            "penyakit kulit", "skin disease", "dermatitis",
        ],
        "Gangguan Pernapasan": [
            "batuk", "cough", "asma", "asthma",
            "gangguan pernapasan", "respiratory",
        ],
        "Luka": [
            "luka", "wound",
        ],
    }

    found = []

    if labelled:
        for item in re.split(
            r",|;|\band\b|\bdan\b",
            labelled,
            flags=re.IGNORECASE,
        ):
            item = re.sub(r"\s+", " ", item).strip(" ,;:.")
            if 2 <= len(item) <= 80:
                found.append(item)

    for canonical, variants in disease_map.items():
        if any(
            re.search(
                rf"(?<!\w){re.escape(variant)}(?!\w)",
                text,
                flags=re.IGNORECASE,
            )
            for variant in variants
        ):
            found.append(canonical)

    return ", ".join(_unique_preserve_order(found)[:8])

def _remove_article_citations(text):
    """Membersihkan sitasi artikel tanpa mengubah angka dosis penting."""
    value = str(text or "")
    value = re.sub(r"\[(?:\d+|\d+\s*[-–]\s*\d+)(?:\s*,\s*\d+)*\]", "", value)
    value = re.sub(
        r"\((?:[A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'’.-]+(?:\s+et\s+al\.)?,?\s*)?"
        r"(?:19|20)\d{2}[a-z]?(?:\s*;\s*[^)]*(?:19|20)\d{2}[a-z]?)*\)",
        "",
        value,
        flags=re.IGNORECASE,
    )
    return re.sub(r"\s+", " ", value).strip(" ,;:.|-\n\t")


def _compact_article_point(text, max_chars=185):
    """Mengubah kalimat artikel menjadi satu poin ringkas dan tetap faktual."""
    value = _remove_article_citations(text)
    value = re.sub(
        r"^(?:cara pengolahan|cara pemakaian|cara pembuatan|preparation|"
        r"processing method|method of preparation|komposisi/dosis|komposisi|"
        r"dosis|dose|dosage|concentration|konsentrasi)\s*[:\-]\s*",
        "",
        value,
        flags=re.IGNORECASE,
    )
    value = re.sub(r"\s+", " ", value).strip(" ,;:.|-")

    if not value:
        return ""

    if len(value) > max_chars:
        truncated = value[:max_chars]
        cut_positions = [
            truncated.rfind("; "),
            truncated.rfind(", "),
            truncated.rfind(" dan "),
            truncated.rfind(" and "),
            truncated.rfind(" "),
        ]
        cut = max(cut_positions)
        if cut >= int(max_chars * 0.60):
            value = truncated[:cut]
        else:
            value = truncated
        value = value.rstrip(" ,;:.") + "…"

    return value


def _extract_section_text(text, headings, max_chars=2200):
    """Mengambil teks setelah judul bagian artikel yang relevan."""
    lines = str(text or "").splitlines()
    headings_lower = [heading.casefold() for heading in headings]
    collected = []

    for index, line in enumerate(lines):
        clean_line = re.sub(r"\s+", " ", line).strip()
        lower = clean_line.casefold()

        if not clean_line or len(clean_line) > 170:
            continue

        if any(heading in lower for heading in headings_lower):
            section_lines = []
            char_count = 0

            for following in lines[index + 1:index + 18]:
                following_clean = re.sub(r"\s+", " ", following).strip()
                if not following_clean:
                    continue

                # Hentikan saat menemukan heading baru yang pendek.
                looks_like_heading = (
                    len(following_clean) <= 90
                    and len(following_clean.split()) <= 10
                    and (
                        following_clean.isupper()
                        or re.match(r"^\d+(?:\.\d+)*\s+[A-ZÀ-ÖØ-Þ]", following_clean)
                    )
                )
                if looks_like_heading and section_lines:
                    break

                section_lines.append(following_clean)
                char_count += len(following_clean)
                if char_count >= max_chars:
                    break

            if section_lines:
                collected.append(" ".join(section_lines))

    return "\n".join(collected)



def _rank_and_summarise_points(candidates, cue_terms, max_points=1):
    """
    Memilih satu nilai paling relevan dari artikel.
    Hasil tidak diberi nomor dan tidak dibuat sebagai poin ringkasan.
    """
    scored = []

    for order, candidate in enumerate(candidates):
        value = _compact_article_point(candidate)
        if not value:
            continue

        value = re.sub(r"^\s*(?:1\)|2\)|3\)|[-•])\s*", "", value)
        lower = value.casefold()

        score = sum(
            2 for cue in cue_terms
            if cue.casefold() in lower
        )
        score += 2 if re.search(
            r"\b\d+(?:[.,]\d+)?\s*"
            r"(?:mg|g|kg|µg|μg|mcg|ml|mL|l|L|%|ppm|"
            r"menit|minute|minutes|jam|hour|hours)\b",
            value,
            re.IGNORECASE,
        ) else 0
        score += 1 if 20 <= len(value) <= 160 else 0
        score -= order * 0.01

        scored.append((score, order, value))

    if not scored:
        return ""

    scored.sort(key=lambda item: (-item[0], item[1]))
    return scored[0][2]




def _extract_preparation(text):
    """
    Mengambil satu cara pengolahan yang disebutkan langsung dalam dokumen.
    Hasil berupa teks pendek, tanpa nomor, tanpa poin, dan tanpa label
    'ringkasan artikel'.
    """
    labels = [
        "cara pengolahan",
        "cara pemakaian",
        "cara pembuatan",
        "preparation",
        "processing method",
        "method of preparation",
        "preparation of extract",
        "extract preparation",
    ]

    keywords = [
        "direbus", "rebus", "diseduh", "seduh",
        "ditumbuk", "tumbuk", "digiling", "dikeringkan",
        "dipotong", "dihaluskan", "diekstraksi",
        "maserasi", "infusa", "dekokta", "decoction",
        "infusion", "boiled", "brewed", "crushed",
        "ground", "dried", "macerated", "extracted",
        "soaked", "powdered", "filtered", "disaring",
        "evaporated", "diuapkan", "heated", "dipanaskan",
    ]

    labelled = _capture_label_value(
        text,
        labels,
        max_length=110,
    )

    if labelled:
        value = _clean_sentence_value(labelled)
        return value[:110].rstrip(" ,;:.-")

    for sentence in _sentences(text):
        lower = sentence.casefold()

        if not any(
            keyword.casefold() in lower
            for keyword in keywords
        ):
            continue

        value = _remove_article_citations(sentence)
        value = _clean_sentence_value(value)

        if len(value) > 110:
            value = value[:110].rsplit(" ", 1)[0]

        return value.rstrip(" ,;:.-")

    return ""



def _extract_composition_dose(text):
    """
    Mengambil satu komposisi, konsentrasi, atau dosis yang tertulis langsung
    dalam dokumen. Hasil singkat, tanpa nomor dan tanpa poin ringkasan.
    """
    labels = [
        "komposisi/dosis",
        "komposisi",
        "dosis",
        "dose",
        "dosage",
        "concentration",
        "konsentrasi",
        "formulation",
        "composition",
        "extract concentration",
        "treatment dose",
    ]

    unit_pattern = re.compile(
        r"\b\d+(?:[.,]\d+)?\s*"
        r"(?:mg|g|kg|µg|μg|mcg|mL|ml|L|%|ppm|"
        r"mol|mM|µM|μM)"
        r"(?:\s*/\s*(?:kg|mL|ml|L))?\b",
        flags=re.IGNORECASE,
    )

    labelled = _capture_label_value(
        text,
        labels,
        max_length=100,
    )

    if labelled:
        value = _clean_sentence_value(labelled)
        return value[:100].rstrip(" ,;:.-")

    for sentence in _sentences(text):
        if not unit_pattern.search(sentence):
            continue

        value = _remove_article_citations(sentence)
        value = _clean_sentence_value(value)

        if len(value) > 100:
            value = value[:100].rsplit(" ", 1)[0]

        return value.rstrip(" ,;:.-")

    return ""


def extract_entities_from_document(document):
    """
    Mengekstrak entitas langsung dari dokumen tanpa menggunakan Excel.

    Hasil mengikuti struktur kolom utama dataset:
    Nama_Tanaman, Nama_Lokal/Daerah, Nama_Latin, Bagian_Tanaman,
    Zat Bioaktif, Khasiat_Efek_Terapeutik, Kategori_Penyakit,
    Komposisi/Dosis, Cara_Pengolahan, Keterangan, dan Sumber_Data.
    """
    text = document.get("text", "")
    missing = "Tidak disebutkan dalam dokumen"

    latin_name = _extract_latin_name(text)

    plant_name, local_name = _extract_plant_and_local_names(
        text,
        latin_name,
        article_title=document.get("title", ""),
    )

    title_plant_name = _extract_plant_name_from_document_context(
        document,
        latin_name,
    )

    if title_plant_name:
        plant_name = title_plant_name

    if not plant_name and latin_name:
        plant_name = latin_name

    if (
        not local_name
        and plant_name
        and clean_text(plant_name) != clean_text(latin_name)
    ):
        local_name = plant_name

    plant_parts = _extract_plant_parts(text)
    compounds = _extract_bioactive_compounds(text)
    activities = _extract_biological_activities(text)
    disease_category = _extract_disease_category(text)
    preparation = _extract_preparation(text)
    composition_dose = _extract_composition_dose(text)

    article_title = _first_nonempty(
        document.get("title"),
        "Judul artikel tidak terdeteksi",
    )
    authors = _first_nonempty(
        document.get("authors"),
        "Penulis tidak terdeteksi",
    )
    year = _first_nonempty(
        document.get("year"),
        "Tahun tidak terdeteksi",
    )

    source_data = (
        f"Judul Artikel: {article_title} | "
        f"Penulis: {authors} | "
        f"Tahun: {year}"
    )

    result = {
        "Nama Tanaman": _first_nonempty(plant_name, missing),
        "Nama Latin": _first_nonempty(latin_name, missing),
        "Nama Lokal/Daerah": _first_nonempty(local_name, missing),
        "Bagian Tanaman": _first_nonempty(plant_parts, missing),
        "Zat Bioaktif": _first_nonempty(compounds, missing),
        "Khasiat/Efek Terapeutik": _first_nonempty(activities, missing),
        "Kategori Penyakit": _first_nonempty(disease_category, missing),
        "Komposisi/Dosis": _first_nonempty(composition_dose, missing),
        "Cara Pengolahan": _first_nonempty(preparation, missing),
        "Keterangan": missing,
        "Sumber Data": source_data,
        "Gambar": "Belum terdeteksi",
        "Judul Artikel": article_title,
        "Penulis Artikel": authors,
        "Tahun Artikel": year,
        "Mode Ekstraksi": "Dokumen langsung tanpa Excel",
    }

    result["Keterangan"] = build_therapeutic_conclusion(result)

    return result



def get_column_map(df):
    """
    Memetakan variasi nama kolom Excel ke atribut internal aplikasi.
    Kandidat pertama mengikuti judul kolom dataset pengguna.
    """
    return {
        "nama": find_col(df, [
            "Nama_Tanaman",
            "Nama Tanaman",
            "Tanaman",
            "Nama Herbal",
            "Nama",
        ]),
        "lokal": find_col(df, [
            "Nama_Lokal/ Daerah",
            "Nama_Lokal/Daerah",
            "Nama Lokal/ Daerah",
            "Nama Lokal/Daerah",
            "Nama Lokal",
            "Nama Daerah",
            "Bahasa Daerah",
            "Bahasa_Daerah",
            "Sinonim",
        ]),
        "latin": find_col(df, [
            "Nama_Latin",
            "Nama Latin",
            "Latin",
            "Scientific Name",
        ]),
        "bagian": find_col(df, [
            "Bagian_Tanaman",
            "Bagian Tanaman",
            "Bagian_Digunakan",
            "Bagian Digunakan",
            "Bagian yang Digunakan",
            "Bagian",
        ]),
        "senyawa": find_col(df, [
            "Zat Bioaktif",
            "Zat_Bioaktif",
            "Senyawa Bioaktif",
            "Senyawa_Bioaktif",
            "Compound",
            "Senyawa",
            "Kandungan",
            "Kandungan Kimia",
            "Komposisi/Kandungan Kimia",
        ]),
        "khasiat": find_col(df, [
            "Khasiat_Efek_Terapeutik",
            "Khasiat/Efek Terapeutik",
            "Khasiat Efek Terapeutik",
            "Khasiat",
            "Manfaat",
            "Benefit",
            "Biological Activity",
            "Biological_Activity",
            "Efek Terapeutik",
        ]),
        "penyakit": find_col(df, [
            "Kategori_Penyakit",
            "Kategori Penyakit",
            "Nama Penyakit",
            "Penyakit",
            "Disease",
        ]),
        "dosis": find_col(df, [
            "Komposisi /Dosis",
            "Komposisi/Dosis",
            "Komposisi_Dosis",
            "Dosis",
            "Komposisi",
            "Dose",
        ]),
        "pengolahan": find_col(df, [
            "Cara_Pengolahan",
            "Cara Pengolahan",
            "Pengolahan",
            "Cara Pemakaian",
            "Preparation",
        ]),
        "keterangan": find_col(df, [
            "Keterangan",
            "Catatan",
            "Deskripsi",
            "Informasi Tambahan",
        ]),
        "sumber": find_col(df, [
            "Sumber_Data",
            "Sumber Data",
            "Sumber",
            "Referensi",
            "Source",
        ]),
        "gambar": find_col(df, [
            "Gambar",
            "Image",
            "Foto",
            "File Gambar",
            "Path Gambar",
            "Nama File Gambar",
        ]),
        "benefit": find_col(df, [
            "Benefit",
            "Manfaat",
            "Deskripsi Manfaat",
        ]),
    }




# =========================================================
# PENCARIAN SINKRON: TANAMAN ↔ PENYAKIT ↔ BIOAKTIF
# =========================================================
RELATION_QUERY_STOPWORDS = {
    "apa", "apakah", "yang", "untuk", "dari", "dengan", "dan", "atau",
    "nama", "tanaman", "herbal", "obat", "bisa", "dapat", "digunakan",
    "gunakan", "membantu", "mengatasi", "mengobati", "terkait", "cocok",
    "rekomendasi", "carikan", "tampilkan", "informasi", "sakit",
    "penyakit", "khasiat", "efek", "terapeutik", "zat", "bioaktif",
    "senyawa", "bagian", "lokal", "daerah", "latin", "sebagai",
    "adalah", "ini", "itu", "saya", "orang", "penderita", "apa",
    "manakah", "mana", "tolong",
}


DISEASE_RELATION_TERMS = {
    "Kanker": {
        "disease": [
            "kanker", "cancer", "tumor", "neoplasma",
        ],
        "activity": [
            "antikanker", "anticancer", "antitumor",
            "sitotoksik", "cytotoxic",
        ],
    },
    "Batuk": {
        "disease": [
            "batuk", "cough",
        ],
        "activity": [
            "antitusif", "antitussive",
            "ekspektoran", "expectorant",
        ],
    },
    "Sakit Kepala": {
        "disease": [
            "sakit kepala", "nyeri kepala", "headache", "migrain",
        ],
        "activity": [
            "analgesik", "analgesic", "antinyeri",
        ],
    },
    "Diare": {
        "disease": [
            "diare", "diarrhea", "disentri",
        ],
        "activity": [
            "antidiare", "antidiarrheal",
        ],
    },
    "Penyakit Jantung": {
        "disease": [
            "penyakit jantung", "gangguan jantung",
            "jantung", "heart disease", "kardiovaskular",
            "cardiovascular",
        ],
        "activity": [
            "kardioprotektif", "cardioprotective",
        ],
    },
    "Penyakit Ginjal": {
        "disease": [
            "penyakit ginjal", "gangguan ginjal", "ginjal",
            "kidney disease", "renal", "batu ginjal",
            "kidney stone",
        ],
        "activity": [
            "nefroprotektif", "nephroprotective",
            "diuretik", "diuretic",
        ],
    },
    "Diabetes": {
        "disease": [
            "diabetes", "diabetes mellitus", "gula darah",
            "hiperglikemia", "hyperglycemia",
        ],
        "activity": [
            "antidiabetes", "antidiabetic",
            "hipoglikemik", "hypoglycemic",
        ],
    },
    "Hipertensi": {
        "disease": [
            "hipertensi", "darah tinggi", "hypertension",
        ],
        "activity": [
            "antihipertensi", "antihypertensive",
        ],
    },
    "Demam": {
        "disease": [
            "demam", "fever",
        ],
        "activity": [
            "antipiretik", "antipyretic",
        ],
    },
    "Peradangan": {
        "disease": [
            "peradangan", "radang", "inflamasi", "inflammation",
        ],
        "activity": [
            "antiinflamasi", "anti inflammatory",
            "anti-inflammatory",
        ],
    },
    "Gangguan Pencernaan": {
        "disease": [
            "gangguan pencernaan", "sakit perut", "nyeri perut",
            "maag", "gastritis", "asam lambung", "kembung",
            "mual", "sembelit", "konstipasi",
        ],
        "activity": [
            "gastroprotektif", "gastroprotective",
            "antiulser", "antiulcer", "laksatif", "laxative",
        ],
    },
    "Infeksi": {
        "disease": [
            "infeksi", "infection", "infeksi bakteri",
            "infeksi jamur",
        ],
        "activity": [
            "antimikroba", "antimicrobial",
            "antibakteri", "antibacterial",
            "antijamur", "antifungal",
        ],
    },
    "Penyakit Hati": {
        "disease": [
            "penyakit hati", "gangguan hati", "liver disease",
            "hepatitis",
        ],
        "activity": [
            "hepatoprotektif", "hepatoprotective",
        ],
    },
    "Kolesterol": {
        "disease": [
            "kolesterol", "cholesterol",
            "hiperlipidemia", "hyperlipidemia",
        ],
        "activity": [
            "antihiperlipidemia", "antihyperlipidemic",
        ],
    },
    "Asam Urat": {
        "disease": [
            "asam urat", "gout",
            "hiperurisemia", "hyperuricemia",
        ],
        "activity": [
            "antihiperurisemia",
        ],
    },
    "Asma": {
        "disease": [
            "asma", "asthma",
        ],
        "activity": [
            "antiasma", "bronkodilator", "bronchodilator",
        ],
    },
    "Flu": {
        "disease": [
            "flu", "influenza", "pilek", "common cold",
        ],
        "activity": [
            "antivirus", "antiviral",
        ],
    },
    "Luka": {
        "disease": [
            "luka", "wound",
        ],
        "activity": [
            "penyembuhan luka", "wound healing", "antiseptik",
        ],
    },
    "Penyakit Kulit": {
        "disease": [
            "penyakit kulit", "dermatitis", "eksim",
            "eczema", "jerawat", "acne",
        ],
        "activity": [
            "dermatoprotektif", "antijerawat",
        ],
    },
}


def _phrase_in_normalized_text(phrase, normalized_text):
    """
    Mencocokkan kata atau frasa secara utuh.
    'hati' tidak akan cocok dengan kata lain yang hanya kebetulan
    memiliki susunan huruf yang sama.
    """
    phrase = clean_text(phrase)
    normalized_text = clean_text(normalized_text)

    if not phrase or not normalized_text:
        return False

    pattern = rf"(?:^|\s){re.escape(phrase)}(?:$|\s)"
    return bool(re.search(pattern, normalized_text))


def _detect_disease_targets(search_text):
    """
    Mendeteksi penyakit/keluhan yang benar-benar ditulis pengguna.
    Pencocokan menggunakan kata/frasa utuh dan memilih istilah terpanjang.
    """
    query = clean_text(search_text)
    detected = []

    for canonical, term_groups in DISEASE_RELATION_TERMS.items():
        aliases = [
            *term_groups.get("disease", []),
            *term_groups.get("activity", []),
        ]

        matched_aliases = [
            clean_text(alias)
            for alias in aliases
            if _phrase_in_normalized_text(alias, query)
        ]

        if matched_aliases:
            longest = max(
                matched_aliases,
                key=lambda value: (len(value.split()), len(value)),
            )
            detected.append(
                {
                    "canonical": canonical,
                    "matched_alias": longest,
                    "disease_terms": [
                        clean_text(value)
                        for value in term_groups.get("disease", [])
                    ],
                    "activity_terms": [
                        clean_text(value)
                        for value in term_groups.get("activity", [])
                    ],
                }
            )

    # Hindari konsep yang tumpang tindih. Contoh:
    # "sakit kepala" tidak boleh berubah menjadi pencarian nyeri umum.
    detected.sort(
        key=lambda item: (
            len(item["matched_alias"].split()),
            len(item["matched_alias"]),
        ),
        reverse=True,
    )

    selected = []
    used_aliases = []

    for item in detected:
        alias = item["matched_alias"]

        if any(
            _phrase_in_normalized_text(alias, existing)
            or _phrase_in_normalized_text(existing, alias)
            for existing in used_aliases
        ):
            continue

        selected.append(item)
        used_aliases.append(alias)

    return selected


def _extract_fallback_query_terms(search_text):
    """
    Mengambil istilah yang belum tersedia dalam kamus penyakit.
    Istilah tetap dicari langsung pada dataset, tetapi hanya sebagai
    kata/frasa utuh sehingga tidak memilih baris secara sembarangan.
    """
    query = clean_text(search_text)

    tokens = [
        token
        for token in query.split()
        if (
            len(token) >= 3
            and token not in RELATION_QUERY_STOPWORDS
        )
    ]

    phrases = []

    # Frasa tiga dan dua kata lebih diprioritaskan daripada satu kata.
    for size in (3, 2):
        for index in range(len(tokens) - size + 1):
            phrases.append(
                " ".join(tokens[index:index + size])
            )

    phrases.extend(tokens)

    return _unique_preserve_order(phrases)


def _prepare_relation_search_dataframe(df, columns):
    """
    Mengisi hanya identitas tanaman yang kosong akibat sel Excel digabung.
    Nilai penyakit, khasiat, senyawa, dosis, dan pengolahan tidak diteruskan
    ke baris lain agar relasinya tidak tercampur dengan tanaman berbeda.
    """
    working = df.copy()

    identity_fields = [
        "nama",
        "lokal",
        "latin",
        "bagian",
        "sumber",
    ]

    for field in identity_fields:
        column = columns.get(field)

        if not column or column not in working.columns:
            continue

        values = (
            working[column]
            .astype(str)
            .replace({
                "": pd.NA,
                "nan": pd.NA,
                "None": pd.NA,
                "NaN": pd.NA,
            })
        )

        if "__sheet_name__" in working.columns:
            values = (
                working
                .assign(__value_to_fill__=values)
                .groupby("__sheet_name__", dropna=False)[
                    "__value_to_fill__"
                ]
                .ffill()
            )
        else:
            values = values.ffill()

        working[column] = values.fillna("")

    return working


def _row_plant_identity(row, columns):
    name = clean_text(
        value_from_row(row, columns.get("nama"))
    )
    latin = clean_text(
        value_from_row(row, columns.get("latin"))
    )

    if latin and latin != "belum terdeteksi":
        return f"latin::{latin}"

    if name and name != "belum terdeteksi":
        return f"name::{name}"

    return ""


def _detect_plant_constraint(working_df, columns, search_text):
    """
    Mendeteksi nama tanaman yang memang ditulis pada pertanyaan.
    Bila pengguna menulis 'jahe untuk batuk', pencarian tidak boleh
    berpindah ke tanaman lain.
    """
    query = clean_text(search_text)
    candidates = []

    for _, row in working_df.iterrows():
        identity = _row_plant_identity(row, columns)

        if not identity:
            continue

        for field in ("nama", "latin", "lokal"):
            column = columns.get(field)

            if not column:
                continue

            value = clean_text(
                value_from_row(row, column)
            )

            if (
                not value
                or value == "belum terdeteksi"
                or len(value) < 3
            ):
                continue

            values = (
                split_multi_value(value)
                if field == "lokal"
                else [value]
            )

            for candidate in values:
                if (
                    candidate
                    and len(candidate) >= 3
                    and _phrase_in_normalized_text(
                        candidate,
                        query,
                    )
                ):
                    candidates.append(
                        {
                            "identity": identity,
                            "value": candidate,
                            "row": row,
                            "length": len(candidate),
                            "words": len(candidate.split()),
                        }
                    )

    if not candidates:
        return None

    candidates.sort(
        key=lambda item: (
            item["words"],
            item["length"],
        ),
        reverse=True,
    )

    return candidates[0]


def _count_complete_fields(row, columns):
    fields = [
        "nama",
        "lokal",
        "latin",
        "bagian",
        "senyawa",
        "khasiat",
        "penyakit",
        "dosis",
        "pengolahan",
        "keterangan",
        "sumber",
    ]

    return sum(
        1
        for field in fields
        if (
            columns.get(field)
            and not _is_missing_value(
                value_from_row(
                    row,
                    columns.get(field),
                )
            )
        )
    )



def _score_known_disease_relation(row, target, columns):
    """
    Menilai relasi penyakit secara ketat.

    ATURAN PENTING:
    - Kategori_Penyakit menjadi sumber utama.
    - Khasiat/Efek Terapeutik menjadi sumber pendukung.
    - Keterangan TIDAK dipakai untuk menentukan penyakit karena bagian
      tersebut dapat berisi peringatan, kontraindikasi, atau efek samping.

    Contoh kesalahan yang dicegah:
    Baris Buncis berkategori Diabetes Mellitus memiliki keterangan
    'hati-hati pada penderita asam urat tinggi'. Baris itu tidak boleh
    dipilih ketika pengguna menanyakan asam urat.
    """
    disease_text = value_from_row(
        row,
        columns.get("penyakit"),
    )
    activity_text = value_from_row(
        row,
        columns.get("khasiat"),
    )

    score = 0
    direct_category_hit = False
    therapeutic_hit = False

    matched_alias = clean_text(
        target.get("matched_alias", "")
    )

    # 1. Kecocokan tepat dengan kata/frasa yang diketik pengguna
    #    pada Kategori_Penyakit mendapat skor tertinggi.
    if (
        matched_alias
        and _phrase_in_normalized_text(
            matched_alias,
            disease_text,
        )
    ):
        score += 12000
        direct_category_hit = True

    # 2. Sinonim penyakit pada Kategori_Penyakit.
    for term in target.get("disease_terms", []):
        if _phrase_in_normalized_text(
            term,
            disease_text,
        ):
            score += 7000
            direct_category_hit = True

    # 3. Aktivitas terapeutik yang benar-benar berhubungan.
    for term in target.get("activity_terms", []):
        if _phrase_in_normalized_text(
            term,
            activity_text,
        ):
            score += 4200
            therapeutic_hit = True

        # Beberapa dataset menempatkan aktivitas pada kolom kategori.
        if _phrase_in_normalized_text(
            term,
            disease_text,
        ):
            score += 3000
            direct_category_hit = True

    # 4. Ada dataset yang menulis nama penyakit langsung pada kolom khasiat.
    for term in target.get("disease_terms", []):
        if _phrase_in_normalized_text(
            term,
            activity_text,
        ):
            score += 2500
            therapeutic_hit = True

    # Baris tanpa kecocokan kategori atau khasiat tidak boleh dipilih.
    if not direct_category_hit and not therapeutic_hit:
        return 0

    # Prioritaskan relasi yang memiliki kategori penyakit langsung.
    if direct_category_hit:
        score += 1800

    return score



def _score_fallback_relation(row, query_terms, columns):
    """
    Pencarian istilah yang belum ada pada kamus.

    Keterangan tidak digunakan sebagai bukti penyakit karena dapat memuat
    larangan, efek samping, atau kalimat 'hati-hati pada penderita ...'.
    """
    field_weights = [
        ("penyakit", 1800),
        ("khasiat", 1300),
        ("senyawa", 700),
        ("nama", 650),
        ("latin", 620),
        ("lokal", 600),
    ]

    score = 0
    found = False

    for field, weight in field_weights:
        column = columns.get(field)

        if not column:
            continue

        value = value_from_row(row, column)

        for position, term in enumerate(query_terms):
            if _phrase_in_normalized_text(term, value):
                score += max(
                    weight - (position * 15),
                    50,
                )
                found = True

    return score if found else 0


def score_row(
    row,
    search_text,
    columns,
    disease_targets=None,
    fallback_terms=None,
):
    """
    Fungsi kompatibilitas untuk penilaian satu baris.
    Skor nol berarti baris tidak memiliki hubungan dengan pertanyaan.
    """
    disease_targets = (
        disease_targets
        if disease_targets is not None
        else _detect_disease_targets(search_text)
    )

    fallback_terms = (
        fallback_terms
        if fallback_terms is not None
        else _extract_fallback_query_terms(search_text)
    )

    if disease_targets:
        relation_score = sum(
            _score_known_disease_relation(
                row,
                target,
                columns,
            )
            for target in disease_targets
        )
    else:
        relation_score = _score_fallback_relation(
            row,
            fallback_terms,
            columns,
        )

    if relation_score <= 0:
        return 0

    # Kelengkapan hanya menjadi pembeda setelah relasi benar ditemukan.
    return (
        relation_score
        + _count_complete_fields(row, columns)
    )


def find_best_match(df, search_text):
    """
    Mencari relasi yang sinkron.

    Aturan:
    1. Bila ada nama tanaman dalam pertanyaan, hasil dibatasi pada tanaman itu.
    2. Bila ada penyakit/keluhan, baris wajib cocok dengan penyakit atau
       efek terapeutik yang berhubungan.
    3. Baris tanpa kecocokan tidak pernah dipilih hanya karena datanya lengkap.
    4. Bila tidak ada relasi yang tepat, sistem menyatakan tidak ditemukan
       dan tidak menampilkan tanaman lain.
    """
    if df.empty:
        return None, "Dataset belum terbaca.", 0

    query = clean_text(search_text)

    if not query:
        return None, "Input tanaman atau kalimat masih kosong.", 0

    columns = get_column_map(df)
    working_df = _prepare_relation_search_dataframe(
        df,
        columns,
    )

    disease_targets = _detect_disease_targets(query)
    plant_constraint = _detect_plant_constraint(
        working_df,
        columns,
        query,
    )

    fallback_terms = (
        []
        if disease_targets
        else _extract_fallback_query_terms(query)
    )

    candidates = []

    for row_index, row in working_df.iterrows():
        # Jangan pindah ke tanaman lain bila pengguna menyebut nama tanaman.
        if plant_constraint:
            current_identity = _row_plant_identity(
                row,
                columns,
            )

            if current_identity != plant_constraint["identity"]:
                continue

        if disease_targets:
            relation_score = sum(
                _score_known_disease_relation(
                    row,
                    target,
                    columns,
                )
                for target in disease_targets
            )
        elif plant_constraint:
            # Pertanyaan hanya menyebut tanaman.
            relation_score = 1000
        else:
            relation_score = _score_fallback_relation(
                row,
                fallback_terms,
                columns,
            )

        if relation_score <= 0:
            continue

        complete_count = _count_complete_fields(
            row,
            columns,
        )

        candidates.append(
            {
                "row": row,
                "row_index": row_index,
                "relation_score": relation_score,
                "complete_count": complete_count,
                "final_score": (
                    relation_score
                    + complete_count
                ),
            }
        )

    if not candidates:
        target_text = (
            ", ".join(
                target["canonical"]
                for target in disease_targets
            )
            if disease_targets
            else ", ".join(fallback_terms[:3])
        )

        if plant_constraint and target_text:
            plant_name = value_from_row(
                plant_constraint["row"],
                columns.get("nama"),
            )
            return (
                None,
                f"{plant_name} ditemukan, tetapi tidak terdapat relasi "
                f"yang sesuai dengan {target_text} pada dataset.",
                0,
            )

        return (
            None,
            f"Tidak ditemukan tanaman yang memiliki relasi tepat dengan "
            f"{target_text or search_text} pada dataset.",
            0,
        )

    candidates.sort(
        key=lambda item: (
            item["relation_score"],
            item["complete_count"],
            -item["row_index"],
        ),
        reverse=True,
    )

    best = candidates[0]
    best_row = best["row"]
    best_score = best["final_score"]

    plant = value_from_row(
        best_row,
        columns.get("nama"),
    )
    latin = value_from_row(
        best_row,
        columns.get("latin"),
    )
    compound = value_from_row(
        best_row,
        columns.get("senyawa"),
    )
    activity = value_from_row(
        best_row,
        columns.get("khasiat"),
    )
    disease = value_from_row(
        best_row,
        columns.get("penyakit"),
    )

    requested_target = (
        ", ".join(
            target["canonical"]
            for target in disease_targets
        )
        if disease_targets
        else (
            plant_constraint["value"]
            if plant_constraint
            else ", ".join(fallback_terms[:3])
        )
    )

    status = (
        f"Pertanyaan: {requested_target} | "
        f"Relasi sinkron: {plant} ({latin}) → "
        f"{compound} → {activity} → {disease} | "
        f"Skor relasi: {best_score}"
    )

    return best_row, status, best_score


def extract_result(row, input_text, df):
    columns = get_column_map(df)

    if row is None:
        return {
            "Nama Tanaman": "Tidak ditemukan",
            "Nama Lokal/Daerah": "Tidak ditemukan",
            "Nama Latin": "Tidak ditemukan",
            "Bagian Tanaman": "Tidak ditemukan",
            "Zat Bioaktif": "Tidak ditemukan",
            "Khasiat/Efek Terapeutik": "Tidak ditemukan",
            "Kategori Penyakit": "Tidak ditemukan",
            "Komposisi/Dosis": "Tidak ditemukan",
            "Cara Pengolahan": "Tidak ditemukan",
            "Keterangan": (
                "Tidak ditemukan relasi tanaman dan penyakit yang "
                "sesuai dengan pertanyaan pada dataset."
            ),
            "Sumber Data": "Dataset HyTBIONEX",
            "Gambar": "Belum terdeteksi",
        }

    result = {
        "Nama Tanaman": value_from_row(row, columns["nama"]),
        "Nama Lokal/Daerah": value_from_row(row, columns["lokal"]),
        "Nama Latin": value_from_row(row, columns["latin"]),
        "Bagian Tanaman": value_from_row(row, columns["bagian"]),
        "Zat Bioaktif": value_from_row(row, columns["senyawa"]),
        "Khasiat/Efek Terapeutik": value_from_row(row, columns["khasiat"]),
        "Kategori Penyakit": value_from_row(row, columns["penyakit"]),
        "Komposisi/Dosis": value_from_row(row, columns["dosis"]),
        "Cara Pengolahan": value_from_row(row, columns["pengolahan"]),
        "Keterangan": value_from_row(row, columns["keterangan"]),
        "Sumber Data": value_from_row(row, columns["sumber"]),
        "Gambar": value_from_row(row, columns["gambar"]),
    }

    result["Keterangan"] = build_therapeutic_conclusion(result)
    return result


def find_plant_image(result):
    """
    Mencari gambar secara toleran berdasarkan:
    1. Isi kolom Gambar.
    2. Nama tanaman.
    3. Nama lokal/daerah.
    4. Nama Latin.
    5. Judul artikel (untuk upload dokumen).

    Pencarian tidak membedakan huruf besar/kecil, spasi, tanda hubung,
    atau garis bawah. Gambar tetap hanya sebagai pendukung tampilan.
    """
    image_value = str(result.get("Gambar", "") or "").strip()

    # Prioritas pertama: path yang ditulis langsung pada dataset.
    if image_value and not _is_missing_value(image_value):
        direct_candidates = [
            Path(image_value),
            Path(ASSET_DIR) / image_value,
            Path("gambar") / image_value,
            Path("images") / image_value,
        ]
        for candidate in direct_candidates:
            if candidate.exists() and candidate.is_file():
                return str(candidate)

    lookup_values = [
        result.get("Nama Tanaman", ""),
        result.get("Nama Lokal/Daerah", ""),
        result.get("Nama Latin", ""),
        result.get("Judul Artikel", ""),
        result.get("Sumber Data", ""),
    ]
    lookup_keys = [
        _normalise_lookup_key(value)
        for value in lookup_values
        if value and not _is_missing_value(value)
    ]
    lookup_keys = [key for key in lookup_keys if key]

    if not lookup_keys:
        return None

    best_path = None
    best_score = 0

    for image_path in _available_plant_images():
        stem_key = _normalise_lookup_key(image_path.stem)
        if not stem_key:
            continue

        stem_tokens = set(stem_key.split())

        for lookup_key in lookup_keys:
            lookup_tokens = set(lookup_key.split())
            score = 0

            if stem_key == lookup_key:
                score = 200
            elif re.search(
                rf"(?<![a-z0-9]){re.escape(stem_key)}(?![a-z0-9])",
                lookup_key,
            ):
                score = 170 + len(stem_key)
            elif re.search(
                rf"(?<![a-z0-9]){re.escape(lookup_key)}(?![a-z0-9])",
                stem_key,
            ):
                score = 150 + len(lookup_key)
            elif stem_tokens and stem_tokens.issubset(lookup_tokens):
                score = 125 + len(stem_tokens) * 5
            else:
                overlap = len(stem_tokens & lookup_tokens)
                union = len(stem_tokens | lookup_tokens)
                similarity = overlap / union if union else 0
                if similarity >= 0.50:
                    score = int(80 + similarity * 40)

            if score > best_score:
                best_score = score
                best_path = image_path

    return str(best_path) if best_path else None



def run_extraction(text_input, uploaded_file, df, dataset_status):
    """
    Dua jalur ekstraksi:
    1. Ada dokumen: entitas diambil langsung dari dokumen, tanpa Excel.
    2. Tanpa dokumen: input tanaman dicocokkan dengan dataset Excel.
    """
    if uploaded_file is not None:
        document, document_status = read_uploaded_document(uploaded_file)

        if document.get("text"):
            result = extract_entities_from_document(document)

            # Input teks pada dashboard digunakan hanya sebagai fallback nama
            # apabila artikel tidak menuliskan nama umum tanaman.
            if (
                text_input.strip()
                and _is_missing_value(result.get("Nama Tanaman"))
            ):
                result["Nama Tanaman"] = text_input.strip()
                result["Keterangan"] = build_therapeutic_conclusion(result)

            extraction_source = "Dokumen langsung — dataset Excel tidak digunakan."
            match_status = (
                "Nama tanaman, nama Latin, nama lokal/daerah, bagian tanaman, "
                "senyawa bioaktif, aktivitas biologis, cara pengolahan, "
                "komposisi/dosis, serta sumber data diekstrak langsung dari artikel."
            )
            score = None
        else:
            source_data = (
                "Judul Artikel: Judul artikel tidak terdeteksi | "
                "Penulis: Penulis tidak terdeteksi | "
                "Tahun: Tahun tidak terdeteksi"
            )
            result = {
                "Nama Tanaman": text_input.strip() or "Tidak terdeteksi",
                "Nama Latin": "Tidak terdeteksi",
                "Nama Lokal/Daerah": "Tidak terdeteksi",
                "Bagian Tanaman": "Tidak terdeteksi",
                "Zat Bioaktif": "Tidak terdeteksi",
                "Khasiat/Efek Terapeutik": "Tidak terdeteksi",
                "Cara Pengolahan": "Tidak terdeteksi",
                "Komposisi/Dosis": "Tidak terdeteksi",
                "Sumber Data": source_data,
                "Kategori Penyakit": "Tidak terdeteksi",
                "Keterangan": "Teks dokumen tidak tersedia untuk diekstrak.",
                "Gambar": "Belum terdeteksi",
                "Mode Ekstraksi": "Dokumen langsung tanpa Excel",
            }
            result["Keterangan"] = build_therapeutic_conclusion(result)
            extraction_source = "Dokumen langsung — dataset Excel tidak digunakan."
            match_status = (
                "Teks dokumen tidak tersedia untuk diekstrak. "
                "PDF kemungkinan berupa scan/gambar."
            )
            score = None

    else:
        document_status = "Tidak ada dokumen yang diunggah."
        row, match_status, score = find_best_match(df, text_input)
        result = extract_result(row, text_input, df)
        result["Mode Ekstraksi"] = "Pencocokan input dengan dataset Excel"
        result["Keterangan"] = build_therapeutic_conclusion(result)
        extraction_source = dataset_status

    image_path = find_plant_image(result)

    st.session_state.last_result = result
    st.session_state.last_image = image_path
    st.session_state.last_status = {
        "dataset": extraction_source,
        "document": document_status,
        "match": match_status,
        "score": score,
    }

    return result, image_path, document_status, match_status


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
    """
    Menampilkan relasi bioaktif lengkap.
    Relasi utama menghubungkan langsung Nama Tanaman dengan
    Kategori Penyakit.
    """
    plant = result.get("Nama Tanaman", "Belum terdeteksi")
    local_name = result.get(
        "Nama Lokal/Daerah",
        "Belum terdeteksi",
    )
    latin = result.get("Nama Latin", "Belum terdeteksi")
    plant_part = result.get(
        "Bagian Tanaman",
        "Belum terdeteksi",
    )
    compound = result.get(
        "Zat Bioaktif",
        "Belum terdeteksi",
    )
    activity = result.get(
        "Khasiat/Efek Terapeutik",
        "Belum terdeteksi",
    )
    disease = result.get(
        "Kategori Penyakit",
        "Belum terdeteksi",
    )
    preparation = result.get(
        "Cara Pengolahan",
        "Belum terdeteksi",
    )
    dose = result.get(
        "Komposisi/Dosis",
        "Belum terdeteksi",
    )
    note = result.get(
        "Keterangan",
        "Belum terdeteksi",
    )
    source = result.get(
        "Sumber Data",
        "Belum terdeteksi",
    )

    relations = [
        [
            plant,
            "digunakan untuk membantu mengatasi",
            disease,
        ],
        [
            plant,
            "mengandung zat bioaktif",
            compound,
        ],
        [
            compound,
            "memiliki khasiat/efek terapeutik",
            activity,
        ],
        [
            activity,
            "terkait dengan kategori penyakit",
            disease,
        ],
        [
            plant,
            "memiliki nama lokal/daerah",
            local_name,
        ],
        [
            plant,
            "memiliki nama Latin",
            latin,
        ],
        [
            plant,
            "menggunakan bagian tanaman",
            plant_part,
        ],
        [
            plant,
            "diolah dengan cara",
            preparation,
        ],
        [
            plant,
            "memiliki komposisi/dosis",
            dose,
        ],
        [
            plant,
            "memiliki keterangan",
            note,
        ],
        [
            plant,
            "didukung oleh sumber data",
            source,
        ],
    ]

    return pd.DataFrame(
        relations,
        columns=[
            "Entitas Sumber",
            "Relasi",
            "Entitas Tujuan",
        ],
    )


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
        "Nama Lokal/Daerah": "lokal",
        "Nama Latin": "latin",
        "Bagian Tanaman": "bagian",
        "Zat Bioaktif": "senyawa",
        "Khasiat/Efek Terapeutik": "khasiat",
        "Kategori Penyakit": "penyakit",
        "Komposisi/Dosis": "dosis",
        "Cara Pengolahan": "pengolahan",
        "Keterangan": "keterangan",
        "Sumber Data": "sumber",
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
            if str(value).strip()
            and str(value).strip().lower() not in {"nan", "none"}
        ]

        unique_values = list(dict.fromkeys(values))
        result[output_key] = (
            " | ".join(unique_values[:5])
            if unique_values
            else "Belum terdeteksi"
        )

    result["Keterangan"] = build_therapeutic_conclusion(result)
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



RECOMMENDATION_TERM_MAP = {
    "sakit perut": [
        "sakit perut", "nyeri perut", "gangguan pencernaan", "pencernaan",
        "maag", "gastritis", "diare", "disentri", "kembung", "mual",
        "gastroprotektif", "gastroprotective", "karminatif", "carminative",
        "antidiare", "antidiarrheal", "antiulser", "antiulcer",
        "antispasmodik", "antispasmodic",
    ],
    "maag": [
        "maag", "gastritis", "asam lambung", "nyeri lambung",
        "gastroprotektif", "gastroprotective", "antiulser", "antiulcer",
    ],
    "diare": [
        "diare", "antidiare", "antidiarrheal", "disentri",
        "gangguan pencernaan",
    ],
    "kembung": [
        "kembung", "karminatif", "carminative", "gangguan pencernaan",
        "mual",
    ],
    "batuk": [
        "batuk", "antitusif", "antitussive", "ekspektoran", "expectorant",
        "pelega tenggorokan",
    ],
    "demam": [
        "demam", "antipiretik", "antipyretic", "penurun panas",
    ],
    "sakit kepala": [
        "sakit kepala", "nyeri kepala", "analgesik", "analgesic",
        "antinyeri",
    ],
    "darah tinggi": [
        "darah tinggi", "hipertensi", "antihipertensi", "antihypertensive",
    ],
    "hipertensi": [
        "hipertensi", "darah tinggi", "antihipertensi", "antihypertensive",
    ],
    "diabetes": [
        "diabetes", "antidiabetes", "antidiabetic", "hipoglikemik",
        "hypoglycemic", "gula darah",
    ],
    "gula darah": [
        "gula darah", "diabetes", "antidiabetes", "antidiabetic",
        "hipoglikemik", "hypoglycemic",
    ],
    "radang": [
        "radang", "inflamasi", "antiinflamasi", "anti inflammatory",
        "anti-inflammatory",
    ],
    "anti inflamasi": [
        "anti inflamasi", "antiinflamasi", "anti inflammatory",
        "anti-inflammatory",
    ],
    "nyeri sendi": [
        "nyeri sendi", "rematik", "asam urat", "analgesik",
        "antiinflamasi",
    ],
    "luka": [
        "luka", "penyembuhan luka", "wound healing", "antiseptik",
        "antimikroba",
    ],
}


def expand_recommendation_terms(keyword):
    """Memperluas keluhan awam menjadi istilah khasiat/penyakit pada dataset."""
    keyword_clean = clean_text(keyword)
    if not keyword_clean:
        return []

    expanded = [keyword_clean]

    for trigger, terms in RECOMMENDATION_TERM_MAP.items():
        trigger_clean = clean_text(trigger)
        if trigger_clean in keyword_clean or keyword_clean in trigger_clean:
            expanded.extend(clean_text(term) for term in terms)

    expanded.extend([
        keyword_clean.replace("anti ", "anti"),
        keyword_clean.replace("-", " "),
    ])

    return _unique_preserve_order(expanded)


def _series_or_default(frame, column, default=""):
    if column and column in frame.columns:
        return (
            frame[column]
            .astype(str)
            .replace({"nan": "", "None": "", "NaN": ""})
            .fillna("")
            .str.strip()
        )

    return pd.Series(
        [default] * len(frame),
        index=frame.index,
        dtype="object",
    )


def _combine_benefit_values(benefit_value, effect_value):
    benefit = str(benefit_value or "").strip()
    effect = str(effect_value or "").strip()
    invalid = {"", "nan", "none", "belum terdeteksi"}

    benefit_ok = benefit.casefold() not in invalid
    effect_ok = effect.casefold() not in invalid

    if benefit_ok and effect_ok:
        if clean_text(benefit) == clean_text(effect):
            return benefit
        return f"{effect} | {benefit}"

    if benefit_ok:
        return benefit

    if effect_ok:
        return effect

    return "Belum terdeteksi"


def _merge_unique_recommendation_values(series, max_items=4):
    values = []

    for value in series:
        value = re.sub(r"\s+", " ", str(value or "")).strip(" |,;")
        if value.casefold() in {
            "", "nan", "none", "belum terdeteksi",
            "tidak disebutkan dalam dokumen",
        }:
            continue
        values.append(value)

    values = _unique_preserve_order(values)

    if not values:
        return "Belum terdeteksi"

    return " | ".join(values[:max_items])


def _prepare_recommendation_dataframe(df, columns):
    """
    Mengisi identitas tanaman yang kosong akibat format Excel bertingkat.
    Pengisian dilakukan per sheet agar nama tanaman tidak tertukar.
    """
    working = df.copy()

    identity_columns = [
        columns.get("nama"),
        columns.get("latin"),
        columns.get("lokal"),
        columns.get("sumber"),
    ]
    identity_columns = [
        column for column in identity_columns
        if column and column in working.columns
    ]

    for column in identity_columns:
        working[column] = (
            working[column]
            .astype(str)
            .replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
        )

        if "__sheet_name__" in working.columns:
            working[column] = (
                working.groupby("__sheet_name__", dropna=False)[column]
                .ffill()
            )
        else:
            working[column] = working[column].ffill()

        working[column] = working[column].fillna("")

    return working



def search_recommendations(df, keyword):
    """
    Mencari tanaman berdasarkan keluhan/penyakit, zat bioaktif, atau khasiat.
    Hasil hanya berisi informasi inti untuk tabel rekomendasi.
    """
    if df.empty or not str(keyword).strip():
        return pd.DataFrame()

    columns = get_column_map(df)
    working = _prepare_recommendation_dataframe(df, columns)
    search_terms = expand_recommendation_terms(keyword)

    if not search_terms:
        return pd.DataFrame()

    search_fields = [
        (columns.get("penyakit"), 10, "Nama Penyakit"),
        (columns.get("khasiat"), 9, "Khasiat"),
        (columns.get("benefit"), 8, "Manfaat"),
        (columns.get("senyawa"), 8, "Zat Bioaktif"),
        (columns.get("nama"), 2, "Nama Tanaman"),
        (columns.get("latin"), 2, "Nama Latin"),
        (columns.get("lokal"), 2, "Nama Lokal/Daerah"),
    ]
    search_fields = [
        item for item in search_fields
        if item[0] and item[0] in working.columns
    ]

    if not search_fields:
        return pd.DataFrame()

    original_keyword = clean_text(keyword)
    scored_rows = []

    for index, row in working.iterrows():
        total_score = 0

        for column, weight, _ in search_fields:
            raw_value = str(row.get(column, "") or "").strip()
            value_clean = clean_text(raw_value)

            if not value_clean:
                continue

            best_score = 0

            for term in search_terms:
                if not term:
                    continue

                if term in value_clean:
                    multiplier = 12 if term == original_keyword else 7
                    current_score = weight * multiplier
                else:
                    term_tokens = {
                        token for token in term.split()
                        if len(token) >= 3
                    }
                    value_tokens = set(value_clean.split())
                    current_score = weight * len(term_tokens & value_tokens)

                best_score = max(best_score, current_score)

            total_score += best_score

        if total_score > 0:
            scored_rows.append({
                "_index": index,
                "Skor Relevansi": total_score,
            })

    if not scored_rows:
        return pd.DataFrame()

    score_df = pd.DataFrame(scored_rows)
    matched = (
        working.loc[score_df["_index"]]
        .copy()
        .reset_index(drop=True)
    )
    matched["Skor Relevansi"] = score_df["Skor Relevansi"].values

    name_series = _series_or_default(matched, columns.get("nama"))
    latin_series = _series_or_default(matched, columns.get("latin"))
    local_series = _series_or_default(matched, columns.get("lokal"))
    compound_series = _series_or_default(matched, columns.get("senyawa"))
    effect_series = _series_or_default(matched, columns.get("khasiat"))
    benefit_series = _series_or_default(matched, columns.get("benefit"))
    disease_series = _series_or_default(matched, columns.get("penyakit"))

    display_name = name_series.mask(
        name_series.str.strip().eq(""),
        local_series,
    )
    display_name = display_name.mask(
        display_name.str.strip().eq(""),
        latin_series,
    ).replace("", "Tidak terdeteksi")

    display_local = local_series.mask(
        local_series.str.strip().eq(""),
        display_name,
    ).replace("", "Tidak terdeteksi")

    display_latin = latin_series.replace("", "Belum terdeteksi")

    effects = [
        _combine_benefit_values(benefit, effect)
        for benefit, effect in zip(benefit_series, effect_series)
    ]

    output = pd.DataFrame({
        "Nama Tanaman": display_name,
        "Nama Latin": display_latin,
        "Nama Lokal/Daerah": display_local,
        "Zat Bioaktif": compound_series.replace(
            "", "Belum terdeteksi"
        ),
        "Khasiat/Efek Terapeutik": effects,
        "Nama Penyakit": disease_series.replace(
            "", "Belum terdeteksi"
        ),
        "Skor Relevansi": matched["Skor Relevansi"],
    })

    output["_plant_key"] = (
        output["Nama Tanaman"].map(clean_text)
        + "|"
        + output["Nama Latin"].map(clean_text)
    )

    grouped_results = []

    for _, group in output.groupby("_plant_key", sort=False):
        grouped_results.append({
            "Nama Tanaman": _merge_unique_recommendation_values(
                group["Nama Tanaman"], max_items=1
            ),
            "Nama Latin": _merge_unique_recommendation_values(
                group["Nama Latin"], max_items=2
            ),
            "Nama Lokal/Daerah": _merge_unique_recommendation_values(
                group["Nama Lokal/Daerah"], max_items=3
            ),
            "Zat Bioaktif": _merge_unique_recommendation_values(
                group["Zat Bioaktif"], max_items=6
            ),
            "Khasiat/Efek Terapeutik": _merge_unique_recommendation_values(
                group["Khasiat/Efek Terapeutik"], max_items=5
            ),
            "Nama Penyakit": _merge_unique_recommendation_values(
                group["Nama Penyakit"], max_items=5
            ),
            "Skor Relevansi": int(group["Skor Relevansi"].max()),
        })

    return (
        pd.DataFrame(grouped_results)
        .sort_values(
            ["Skor Relevansi", "Nama Tanaman"],
            ascending=[False, True],
        )
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

css_styles = """<style>
.stApp {
    background: #f5fbf7;
    font-family: "Poppins", "Segoe UI", "Trebuchet MS", Arial, sans-serif;
}
.main .block-container {
    padding-top: 1rem;
    max-width: 1500px;
}
[data-testid="stSidebar"] {
    background:
        radial-gradient(circle at bottom left, rgba(34,197,94,0.25), transparent 28%),
        radial-gradient(circle at top right, rgba(16,185,129,0.18), transparent 35%),
        linear-gradient(180deg, #021f16 0%, #043b2c 45%, #065f46 100%);
    border-right: 1px solid rgba(255,255,255,0.12);
}
[data-testid="stSidebar"] * {
    color: #f8fff8;
}
[data-testid="stSidebar"] .stButton > button {
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
}
[data-testid="stSidebar"] .stButton > button:hover {
    background: rgba(255,255,255,0.12);
    color: #ffffff;
}
.nav-active {
    background: linear-gradient(90deg, rgba(34,197,94,0.38), rgba(16,185,129,0.20));
    border: 1px solid rgba(134,239,172,0.28);
    border-radius: 0.8rem;
    padding: 0.78rem 0.85rem;
    margin-bottom: 0.3rem;
    color: #ffffff;
    font-weight: 800;
}
.sidebar-section-title {
    color: #86efac;
    font-size: 0.75rem;
    font-weight: 900;
    letter-spacing: 0.08rem;
    margin-top: 1.25rem;
    margin-bottom: 0.45rem;
}
.sidebar-line {
    height: 1px;
    background: rgba(255,255,255,0.14);
    margin: 1rem 0;
}
.top-header {
    background: linear-gradient(135deg, #013220, #064e3b, #065f46);
    padding: 1.4rem 1.7rem;
    border-radius: 0 0 1.6rem 1.6rem;
    color: white;
    margin-bottom: 1.25rem;
    box-shadow: 0 16px 38px rgba(0,0,0,0.18);
}
.top-header h1 {
    margin: 0;
    color: white;
    font-size: 2.1rem;
    font-weight: 900;
}
.top-header p {
    margin: 0.35rem 0 0;
    color: #d1fae5;
}
.hero-banner {
    background: __HERO_BACKGROUND__;
    background-size: cover;
    background-position: center;
    padding: 2rem;
    border-radius: 1.35rem;
    min-height: 220px;
    box-shadow: inset 0 0 0 1px rgba(6,78,59,0.08);
}
.hero-banner h2 {
    color: #047857;
    font-size: 1.85rem;
    font-weight: 900;
    margin-bottom: 0.8rem;
}
.hero-banner p {
    color: #12372a;
    font-size: 1rem;
    line-height: 1.65;
    max-width: 700px;
}
.section-title {
    color: #064e3b;
    font-size: 1.6rem;
    font-weight: 900;
    letter-spacing: -0.02em;
    margin-top: 1.25rem;
    margin-bottom: 0.85rem;
    font-family: "Poppins", "Segoe UI", "Trebuchet MS", Arial, sans-serif;
}
[data-testid="stMetric"] {
    background: white;
    border: 1px solid rgba(6,78,59,0.10);
    border-radius: 1rem;
    padding: 1rem;
    box-shadow: 0 8px 20px rgba(15,23,42,0.07);
}
[data-testid="stMetricLabel"] {
    color: #065f46;
    font-weight: 800;
}
[data-testid="stMetricValue"] {
    color: #047857;
    font-weight: 900;
}
.result-card {
    background: linear-gradient(145deg, #ffffff 0%, #fbfffd 100%);
    border-left: 0.42rem solid #059669;
    border-radius: 1rem;
    padding: 1.05rem 1.1rem;
    min-height: 122px;
    box-shadow: 0 10px 24px rgba(15,23,42,0.09);
    margin-bottom: 0.8rem;
    transition: transform 0.18s ease, box-shadow 0.18s ease;
    font-family: "Poppins", "Segoe UI", "Trebuchet MS", Arial, sans-serif;
}
.result-card:hover {
    transform: translateY(-2px);
    box-shadow: 0 14px 30px rgba(15,23,42,0.13);
}
.result-card-title {
    color: #065f46;
    font-size: 1.02rem;
    font-weight: 850;
    line-height: 1.28;
    letter-spacing: 0.01em;
    margin-bottom: 0.72rem;
}
.result-card-value {
    color: #1f2937;
    font-size: 0.96rem;
    font-weight: 500;
    line-height: 1.55;
    overflow-wrap: anywhere;
    word-break: break-word;
}
.result-card.bioactive-card,
.result-card.activity-card {
    background: linear-gradient(145deg, #fffaf0 0%, #fff7e6 100%);
    border: 2px solid #fdba74;
    border-left: 0.52rem solid #fb923c;
}
.result-card.bioactive-card .result-card-title,
.result-card.activity-card .result-card-title {
    color: #9a3412;
    font-weight: 900;
}
.result-card.bioactive-card .result-card-value,
.result-card.activity-card .result-card-value {
    color: #7c2d12;
    font-weight: 600;
}
.result-card.disease-card {
    background: linear-gradient(145deg, #fef2f2 0%, #fff7f7 100%);
    border-left-color: #e11d48;
}
.result-card.disease-card .result-card-title {
    color: #9f1239;
}
.result-card.note-card {
    background: linear-gradient(145deg, #f0fdf4 0%, #f8fffa 100%);
    border-left-color: #16a34a;
}
.result-card.note-card .result-card-value {
    font-size: 1rem;
    font-weight: 550;
}
.info-box {
    background: #f3e8ff;
    padding: 1rem;
    border-radius: 1rem;
    border: 1px solid #c084fc;
    color: #111827;
    margin-bottom: 1rem;
}
.stButton > button {
    background: linear-gradient(90deg, #047857, #059669);
    color: white;
    border: none;
    border-radius: 0.8rem;
    font-weight: 900;
}
.stButton > button:hover {
    background: linear-gradient(90deg, #065f46, #047857);
    color: white;
}
[data-testid="stFileUploader"] section {
    background: #fbfffb;
    border: 2px dashed #86efac;
    border-radius: 1rem;
}
textarea {
    background: white !important;
    color: #0f172a !important;
}
.quick-card {
    background: white;
    border-radius: 1rem;
    padding: 1rem;
    border: 1px solid rgba(6,78,59,0.10);
    box-shadow: 0 8px 20px rgba(15,23,42,0.07);
    min-height: 160px;
    text-align: center;
}
.quick-card h4 {
    color: #064e3b;
    margin: 0.4rem 0;
}
.quick-card p {
    color: #475569;
    font-size: 0.88rem;
}
.bioactive-main-card {
    background:
        radial-gradient(circle at top right, rgba(254,215,170,0.78), transparent 34%),
        linear-gradient(135deg, #fff7ed 0%, #fffbeb 52%, #ecfdf5 100%);
    border: 2px solid #f59e0b;
    border-left: 0.75rem solid #f97316;
    border-radius: 1.3rem;
    padding: 1.45rem;
    margin-top: 1rem;
    margin-bottom: 1.25rem;
    box-shadow: 0 14px 32px rgba(249,115,22,0.15);
}
.bioactive-main-card h2 {
    color: #9a3412;
    font-size: 1.65rem;
    font-weight: 900;
    margin: 0 0 0.4rem 0;
}
.bioactive-number {
    color: #ea580c;
    font-size: 2.9rem;
    font-weight: 900;
    line-height: 1;
    margin-top: 0.55rem;
}
.bioactive-description {
    color: #475569;
    font-size: 0.98rem;
    line-height: 1.6;
    margin-top: 0.7rem;
}
.bioactive-chip {
    display: inline-block;
    background: #ffffff;
    color: #9a3412;
    border: 1px solid #fdba74;
    border-radius: 999px;
    padding: 0.38rem 0.7rem;
    margin: 0.35rem 0.25rem 0 0;
    font-size: 0.82rem;
    font-weight: 800;
}
.bioactive-result {
    background: linear-gradient(135deg, #fff7ed 0%, #fffbeb 55%, #f0fdf4 100%);
    border: 2px solid #fb923c;
    border-left: 0.75rem solid #f97316;
    border-radius: 1.15rem;
    padding: 1.3rem;
    margin-top: 1rem;
    margin-bottom: 1.15rem;
    box-shadow: 0 12px 28px rgba(249,115,22,0.14);
}
.bioactive-result h2 {
    color: #9a3412;
    font-size: 1.45rem;
    font-weight: 900;
    margin: 0 0 0.7rem 0;
}
.bioactive-compound {
    color: #c2410c;
    font-size: 1.6rem;
    font-weight: 900;
    margin-bottom: 0.65rem;
    overflow-wrap: anywhere;
}
.bioactive-relation {
    color: #065f46;
    font-size: 1rem;
    font-weight: 700;
    line-height: 1.65;
    overflow-wrap: anywhere;
}
.bioactive-summary-text {
    color: #1f2937;
    font-size: 1.05rem;
    line-height: 1.75;
    margin-top: 0.8rem;
    padding: 1rem;
    background: rgba(255,255,255,0.78);
    border: 1px solid #fed7aa;
    border-radius: 0.9rem;
}
.bioactive-source-text {
    color: #475569;
    font-size: 0.9rem;
    line-height: 1.55;
    margin-top: 0.8rem;
}

</style>"""

# Masukkan background dinamis tanpa memakai f-string pada CSS.
css_styles = css_styles.replace(
    "__HERO_BACKGROUND__",
    hero_background_css,
)

st.markdown(
    css_styles,
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
    "📋 Bioaktif Informasi Ekstraksi",
    "📋 Bioaktif Informasi Ekstraksi",
    "nav_entity_result",
)
sidebar_nav_button(
    "🔗 Relasi Ekstraksi Bioaktif",
    "🔗 Relasi Ekstraksi",
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
# STATISTIK SENYAWA BIOAKTIF — SATU SUMBER PERHITUNGAN
# =========================================================
def get_bioactive_statistics(df, top_n=10):
    """Menghitung senyawa bioaktif unik dan frekuensi kemunculannya."""
    compound_col = get_column_map(df).get("senyawa")

    if df.empty or compound_col is None:
        return 0, pd.DataFrame(columns=["Senyawa Bioaktif", "Jumlah Kemunculan"])

    counter = Counter()
    display_names = {}

    for value in df[compound_col].tolist():
        if pd.isna(value):
            continue

        raw_value = str(value).strip()
        if not raw_value or raw_value.lower() in {"nan", "none", "belum terdeteksi"}:
            continue

        for compound in re.split(r"[,;/|]+", raw_value):
            compound = re.sub(r"\s+", " ", compound).strip()
            if not compound:
                continue

            key = compound.casefold()
            counter[key] += 1
            display_names.setdefault(key, compound)

    top_rows = [
        {
            "Senyawa Bioaktif": display_names[key],
            "Jumlah Kemunculan": count,
        }
        for key, count in counter.most_common(top_n)
    ]

    return len(counter), pd.DataFrame(top_rows)


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
# Angka ini dipakai bersama oleh kartu metrik atas dan kartu bioaktif utama.
# Setiap sel dipisahkan menjadi senyawa individual, lalu dihitung unik
# tanpa membedakan huruf besar dan kecil.
total_compounds, top_compounds_dashboard = get_bioactive_statistics(
    df_data,
    top_n=10,
)
total_effects = (
    df_data[columns["khasiat"]].astype(str).replace("", pd.NA).dropna().nunique()
    if columns["khasiat"] else 0
)
total_relations = total_rows * 10 if total_rows else 0

# =========================================================
# KOMPONEN TAMPILAN
# =========================================================
def render_metrics():
    metric_cols = st.columns(5)
    metric_cols[0].metric("Total Baris", f"{total_rows:,}")
    metric_cols[1].metric("Total Tanaman", f"{total_plants:,}")
    metric_cols[2].metric("Senyawa Bioaktif Unik", f"{total_compounds:,}")
    metric_cols[3].metric("Aktivitas Biologis", f"{total_effects:,}")
    metric_cols[4].metric("Relasi Triplet", f"{total_relations:,}")



def render_bioactive_dashboard():
    """
    Menonjolkan senyawa bioaktif sebagai fokus utama dashboard.

    Angka pada kartu ini berasal dari variabel yang sama dengan metrik
    "Senyawa Bioaktif Unik" di bagian atas, sehingga nilainya selalu sama.
    """
    total_unique = total_compounds
    top_compounds = top_compounds_dashboard.copy()

    if not top_compounds.empty:
        chips = "".join(
            f'<span class="bioactive-chip">🧪 {safe_text(row["Senyawa Bioaktif"])}</span>'
            for _, row in top_compounds.head(7).iterrows()
        )
    else:
        chips = '<span class="bioactive-chip">Data senyawa belum ditemukan</span>'

    st.markdown(
        f"""<div class="bioactive-main-card">
<h2>🧪  Kandungan Zat Bioaktif Tanaman Herbal</h2>
<div class="bioactive-number">{total_unique:,}</div>
<div class="bioactive-description">
Total senyawa bioaktif unik yang teridentifikasi pada dataset HyTBIONEX.
Isi setiap sel dipisahkan menjadi senyawa individual, dibersihkan,
lalu dideduplikasi tanpa membedakan huruf besar dan kecil.
Senyawa bioaktif menjadi penghubung utama antara tanaman herbal,
bagian tanaman, aktivitas biologis/efek terapeutik, dan sumber bukti.
</div>
<div style="margin-top:0.65rem;">{chips}</div>
</div>""",
        unsafe_allow_html=True,
    )

    if not top_compounds.empty:
        chart_df = top_compounds.sort_values("Jumlah Kemunculan", ascending=True)
        fig = px.bar(
            chart_df,
            x="Jumlah Kemunculan",
            y="Senyawa Bioaktif",
            orientation="h",
            text="Jumlah Kemunculan",
            title="Senyawa Bioaktif yang Paling Sering Ditemukan",
        )
        fig.update_traces(textposition="outside")
        fig.update_layout(
            height=430,
            plot_bgcolor="white",
            paper_bgcolor="white",
            xaxis_title="Jumlah Kemunculan",
            yaxis_title="Senyawa Bioaktif",
            margin=dict(l=20, r=35, t=65, b=20),
        )
        st.plotly_chart(
            fig,
            use_container_width=True,
            key="dashboard_bioactive_chart",
            config={"displaylogo": False},
        )




def render_bioactive_focus(result):
    """
    Menampilkan Zat Bioaktif sebagai output utama,
    disertai Keterangan dalam satu kalimat langsung.
    """
    plant = result.get(
        "Nama Tanaman",
        "Belum terdeteksi",
    )
    compound = result.get(
        "Zat Bioaktif",
        "Belum terdeteksi",
    )
    activity = result.get(
        "Khasiat/Efek Terapeutik",
        "Belum terdeteksi",
    )
    disease = result.get(
        "Kategori Penyakit",
        "Belum terdeteksi",
    )
    conclusion = build_therapeutic_conclusion(result)
    source = result.get(
        "Sumber Data",
        "Belum terdeteksi",
    )

    st.markdown(
        f"""<div class="bioactive-result">
<h2>🧪 OUTPUT UTAMA: ZAT BIOAKTIF</h2>
<div class="bioactive-compound">{safe_text(compound)}</div>
<div class="bioactive-summary-text">
<b>Nama Tanaman:</b> {safe_text(plant)}<br>
<b>Khasiat:</b> {safe_text(activity)}<br>
<b>Kategori Penyakit:</b> {safe_text(disease)}<br><br>
<b>Keterangan:</b> {safe_text(conclusion)}
</div>
<div class="bioactive-source-text">
📚 <b>Sumber_Data:</b> {safe_text(source)}
</div>
</div>""",
        unsafe_allow_html=True,
    )


def render_status_box(dataset_text, document_text, match_text):
    st.markdown(
        f"""<div class="info-box">
<b>Mode/Sumber Ekstraksi:</b> {safe_text(dataset_text)}<br>
<b>Status Dokumen:</b> {safe_text(document_text)}<br>
<b>Keterangan Ekstraksi:</b> {safe_text(match_text)}
</div>""",
        unsafe_allow_html=True,
    )


def render_analysis_form(prefix, allow_text=True, allow_upload=True):
    st.markdown(
        '<div class="section-title">📝 Input dan Proses Ekstraksi</div>',
        unsafe_allow_html=True,
    )

    if allow_upload:
        st.caption(
            "Saat dokumen diunggah, seluruh entitas diambil langsung dari isi artikel. "
            "Dataset Excel tidak digunakan pada mode dokumen."
        )

    text_input = ""

    if allow_text and allow_upload:
        left, right = st.columns([1.1, 1])

        with left:
            text_input = st.text_area(
                "Nama tanaman atau kalimat",
                placeholder=(
                    "Contoh: Apa nama tanaman herbal untuk batuk? "
                    "Atau: tanaman untuk diare, ginjal, jantung, dan lainnya."
                ),
                height=145,
                key=f"{prefix}_text",
            )

        with right:
            uploaded_file = st.file_uploader(
                "Unggah dokumen",
                type=["pdf", "docx", "txt", "csv", "xlsx", "xls"],
                key=f"{prefix}_upload",
            )

    elif allow_text:
        text_input = st.text_area(
            "Nama tanaman atau kalimat",
            placeholder=(
                "Contoh: Apa nama tanaman herbal untuk batuk? "
                "Atau ketik penyakit, nama tanaman, khasiat, atau zat bioaktif."
            ),
            height=150,
            key=f"{prefix}_text",
        )
        uploaded_file = None

    else:
        uploaded_file = st.file_uploader(
            "Unggah dokumen",
            type=["pdf", "docx", "txt", "csv", "xlsx", "xls"],
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
            dataset_text=st.session_state.last_status.get("dataset", dataset_status),
            document_text=document_status,
            match_text=match_status,
            key_prefix=f"{prefix}_output",
        )






def render_result_cards(result):
    """
    Menampilkan hasil ekstraksi dalam bentuk kartu hijau-putih dengan
    Zat Bioaktif dan Aktivitas Biologis sebagai fokus berwarna oranye.

    Susunan:
    Baris 1:
    - Senyawa Bioaktif
    - Aktivitas Biologis / Efek Terapeutik
    - Nama Tanaman

    Baris 2:
    - Nama Latin
    - Nama Lokal/Daerah
    - Bagian Tanaman

    Baris 3:
    - Cara Pengolahan
    - Komposisi/Dosis
    - Sumber Data

    Baris 4:
    - Kategori Penyakit
    - Keterangan
    """

    st.markdown(
        '<div class="section-title">'
        '📋 Bioaktif Informasi Ekstraksi'
        '</div>',
        unsafe_allow_html=True,
    )

    def show_card(column, title, value, icon, extra_class=""):
        value = value if value not in (None, "") else "Belum terdeteksi"

        with column:
            st.markdown(
                f"""<div class="result-card {extra_class}">
<div class="result-card-title">{icon} {safe_text(title)}</div>
<div class="result-card-value">{safe_text(value)}</div>
</div>""",
                unsafe_allow_html=True,
            )

    # Baris 1: fokus utama bioaktif.
    row_1 = st.columns(3)

    show_card(
        row_1[0],
        "SENYAWA BIOAKTIF",
        result.get("Zat Bioaktif", "Belum terdeteksi"),
        "🧪",
        "bioactive-card",
    )
    show_card(
        row_1[1],
        "AKTIVITAS BIOLOGIS / EFEK TERAPEUTIK",
        result.get(
            "Khasiat/Efek Terapeutik",
            "Belum terdeteksi",
        ),
        "💚",
        "activity-card",
    )
    show_card(
        row_1[2],
        "Nama Tanaman",
        result.get("Nama Tanaman", "Belum terdeteksi"),
        "🌿",
    )

    # Baris 2: identitas tanaman.
    row_2 = st.columns(3)

    show_card(
        row_2[0],
        "Nama Latin",
        result.get("Nama Latin", "Belum terdeteksi"),
        "🔬",
    )
    show_card(
        row_2[1],
        "Nama Lokal/Daerah",
        result.get("Nama Lokal/Daerah", "Belum terdeteksi"),
        "🇮🇩",
    )
    show_card(
        row_2[2],
        "Bagian Tanaman",
        result.get("Bagian Tanaman", "Belum terdeteksi"),
        "🍃",
    )

    # Baris 3: penggunaan dan sumber.
    row_3 = st.columns(3)

    show_card(
        row_3[0],
        "Cara Pengolahan",
        result.get("Cara Pengolahan", "Belum terdeteksi"),
        "☕",
    )
    show_card(
        row_3[1],
        "Komposisi/Dosis",
        result.get("Komposisi/Dosis", "Belum terdeteksi"),
        "⚖️",
    )
    show_card(
        row_3[2],
        "Sumber Data",
        result.get("Sumber Data", "Belum terdeteksi"),
        "📚",
    )

    # Baris 4: tambahan sesuai permintaan.
    row_4 = st.columns([1, 2])

    show_card(
        row_4[0],
        "Kategori Penyakit",
        result.get("Kategori Penyakit", "Belum terdeteksi"),
        "🩺",
        "disease-card",
    )
    show_card(
        row_4[1],
        "Keterangan",
        result.get("Keterangan", "Belum terdeteksi"),
        "📝",
        "note-card",
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
                "Gambar pendukung ditemukan dari folder assets berdasarkan nama tanaman "
                "hasil ekstraksi. Gambar tidak digunakan sebagai sumber entitas."
            )
            st.write("**Nama Tanaman:**", result.get("Nama Tanaman", ""))
            st.write("**Nama Latin:**", result.get("Nama Latin", ""))
    else:
        plant_name = result.get("Nama Tanaman", "tanaman")
        suggested_name = slugify_filename(plant_name) or "nama_tanaman"
        st.warning(
            "Gambar belum ditemukan. Pastikan file gambar sudah diunggah ke GitHub "
            f"pada folder assets, misalnya assets/{suggested_name}.jpg. "
            "Nama file boleh menggunakan huruf kecil, spasi, tanda hubung, atau garis bawah."
        )


def render_relation_table(result, key_prefix):
    st.markdown(
        '<div class="section-title">🔗 Bioactive Relation Extraction</div>',
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
    render_bioactive_focus(result)
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
            "Kategori Penyakit": "Infeksi dan peradangan",
            "Keterangan": "Serai mengandung citronellal dan geraniol yang berkaitan dengan aktivitas antimikroba.",
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
            "Bioaktif Informasi Ekstraksi",
            "Lihat kembali entitas hasil ekstraksi informasi bioaktif.",
            "📋 Bioaktif Informasi Ekstraksi",
            "quick_entity",
        ),
        (
            "🔗",
            "Bioactive Relation Extraction",
            "Lihat relasi utama tanaman, senyawa bioaktif, aktivitas biologis, dan sumber bukti.",
            "🔗 Relation Extraction",
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
        render_bioactive_focus(result)
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

    st.write(
        "Masukkan keluhan atau penyakit, zat bioaktif, maupun khasiat. "
        "Contoh: **sakit perut**, **batuk**, **hipertensi**, "
        "**flavonoid**, atau **antiinflamasi**."
    )

    keyword = st.text_input(
        "Keluhan, penyakit, zat bioaktif, atau khasiat",
        placeholder=(
            "Contoh: sakit perut, maag, batuk, diabetes, flavonoid, "
            "antiinflamasi"
        ),
        key="recommendation_keyword",
    )

    if st.button(
        "💡 Cari Rekomendasi",
        use_container_width=True,
        key="recommendation_run",
    ):
        if not keyword.strip():
            st.warning("Masukkan kata pencarian terlebih dahulu.")
            return

        recommendations = search_recommendations(df_data, keyword)

        if recommendations.empty:
            st.warning(
                "Belum ditemukan tanaman yang sesuai. "
                "Coba gunakan istilah lain yang terdapat pada dataset."
            )
            return

        display_columns = [
            "Nama Tanaman",
            "Nama Latin",
            "Nama Lokal/Daerah",
            "Zat Bioaktif",
            "Khasiat/Efek Terapeutik",
            "Nama Penyakit",
        ]

        st.success(
            f"Ditemukan {len(recommendations):,} tanaman herbal yang terkait."
        )

        st.dataframe(
            recommendations[display_columns].head(30),
            use_container_width=True,
            hide_index=True,
            height=min(800, 105 + 62 * min(len(recommendations), 30)),
            column_config={
                "Nama Tanaman": st.column_config.TextColumn(
                    "Nama Tanaman",
                    width="medium",
                ),
                "Nama Latin": st.column_config.TextColumn(
                    "Nama Latin",
                    width="medium",
                ),
                "Nama Lokal/Daerah": st.column_config.TextColumn(
                    "Nama Lokal/Daerah",
                    width="medium",
                ),
                "Zat Bioaktif": st.column_config.TextColumn(
                    "Zat Bioaktif",
                    width="large",
                ),
                "Khasiat/Efek Terapeutik": st.column_config.TextColumn(
                    "Khasiat/Efek Terapeutik",
                    width="large",
                ),
                "Nama Penyakit": st.column_config.TextColumn(
                    "Nama Penyakit",
                    width="large",
                ),
            },
        )

        st.caption(
            "Tabel ini menampilkan kecocokan berdasarkan data yang tersedia "
            "dalam dataset HyTBIONEX."
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
        **HyTBIONEX** HyTBIONEX merupakan prototipe sistem cerdas untuk mengekstraksi, 
        mengintegrasikan, dan memvisualisasikan informasi bioaktif tanaman herbal Indonesia.
        Sistem ini menghubungkan entitas nama tanaman, nama Latin, nama lokal atau daerah, 
        bagian tanaman, senyawa bioaktif, aktivitas biologis atau efek terapeutik, cara pengolahan, 
        dosis atau komposisi, serta sumber data ke dalam HerbKG 2.0. Pada input dokumen,
        sumber data diperoleh secara langsung dari judul artikel, nama penulis,  dan tahun publikasi, 
        sehingga setiap informasi yang diekstraksi dapat ditelusuri kembali secara terstruktur dan berbasis bukti ilmiah.

        **Pipeline:** Input → Preprocessing → Adaptive Fine Tuning → BIE → NED → Hybrid Transformer → BRE 
        → HerbKG 2.0 → Aplikasi Downstream.

        **Peneliti:** NAZWITA
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
Platform cerdas untuk ekstraksi kandungan bioaktif tanaman herbal Indonesia,
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
        st.caption("Adaptive Fine Tuning → BIE → NED → Hybrid Transformer → BRE → HerbKG 2.0")

    render_metrics()
    render_bioactive_dashboard()
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

elif page == "📋 Bioaktif Informasi Ekstraksi":
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
        render_bioactive_focus(st.session_state.last_result)
        render_result_cards(st.session_state.last_result)
        render_image_section(
            st.session_state.last_result,
            st.session_state.last_image,
        )
    else:
        st.warning(
            "Belum ada hasil ekstraksi. Jalankan Proses Ekstraksi terlebih dahulu."
        )

elif page == "🔗 Relasi Ekstraksi":
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
