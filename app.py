import streamlit as st

# ==============================
# KONFIGURASI HALAMAN
# ==============================
st.set_page_config(
    page_title="HyTBIONEX",
    page_icon="🌿",
    layout="wide"
)

# ==============================
# CSS TAMPILAN
# ==============================
st.markdown("""
<style>
.stApp {
    background: linear-gradient(135deg, #e8f8ec 0%, #f3e8ff 100%);
}

.main-title {
    background: linear-gradient(135deg, #0b7a45, #129157);
    padding: 30px;
    border-radius: 25px;
    text-align: center;
    color: white;
    margin-bottom: 25px;
    box-shadow: 0 8px 20px rgba(0,0,0,0.18);
}

.main-title h1 {
    font-size: 48px;
    font-weight: 900;
    margin-bottom: 5px;
}

.main-title h3 {
    font-size: 22px;
    font-weight: 700;
}

.card-orange {
    background: linear-gradient(135deg, #f97316, #fb923c);
    padding: 22px;
    border-radius: 20px;
    color: white;
    margin-bottom: 20px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.15);
}

.card-green {
    background: linear-gradient(135deg, #0b7a45, #129157);
    padding: 22px;
    border-radius: 20px;
    color: white;
    margin-bottom: 20px;
    min-height: 280px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.15);
}

.card-lilac {
    background: #f3e8ff;
    padding: 20px;
    border-radius: 18px;
    border: 2px solid #c084fc;
    color: #111111;
    margin-bottom: 18px;
}

.result-card {
    background: white;
    padding: 22px;
    border-radius: 20px;
    border-left: 8px solid #0b7a45;
    color: #111111;
    margin-top: 20px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.12);
}

.sidebar-title {
    font-size: 28px;
    font-weight: 900;
    color: #0b7a45;
}

</style>
""", unsafe_allow_html=True)

# ==============================
# SIDEBAR
# ==============================
st.sidebar.markdown('<div class="sidebar-title">🌱 HyTBIONEX</div>', unsafe_allow_html=True)
st.sidebar.markdown("---")
st.sidebar.markdown("🏠 Dashboard Utama")
st.sidebar.markdown("🌿 Input Data Tanaman")
st.sidebar.markdown("📁 Upload Dokumen")
st.sidebar.markdown("📋 Hasil Ekstraksi")
st.sidebar.markdown("🔗 Relation Extraction")
st.sidebar.markdown("🕸️ HerbKG 2.0")
st.sidebar.markdown("---")
st.sidebar.markdown("**Advanced Downstream Applications**")
st.sidebar.markdown("📊 Descriptive Analytics")
st.sidebar.markdown("🔎 Evidence-Based Graph Query")
st.sidebar.markdown("🧬 Similarity Analysis")
st.sidebar.markdown("💊 Herbal Recommendation")

# ==============================
# HEADER
# ==============================
st.markdown("""
<div class="main-title">
    <h1>🌿 HyTBIONEX</h1>
    <h3>Hybrid Transformer for Bioactive Information Extraction</h3>
    <p>Analisis Bioaktif Tanaman Herbal Indonesia & Enhanced Herb Knowledge Graph 2.0</p>
</div>
""", unsafe_allow_html=True)

# ==============================
# WELCOME
# ==============================
st.markdown("""
<div class="card-orange">
    <h2>Selamat Datang di HyTBIONEX</h2>
    <p>
    Platform cerdas untuk ekstraksi informasi bioaktif tanaman herbal Indonesia
    berbasis Hybrid Transformer, Bioactive Information Extraction, Named Entity Disambiguation,
    Relation Extraction, dan Enhanced Herb Knowledge Graph.
    </p>
</div>
""", unsafe_allow_html=True)

# ==============================
# INPUT AREA
# ==============================
col1, col2 = st.columns(2)

with col1:
    st.markdown("""
    <div class="card-green">
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
    <div class="card-green">
        <h2>📁 2. Upload Dokumen Artikel / Dataset</h2>
        <p>Upload PDF, TXT, CSV, atau Excel.</p>
    </div>
    """, unsafe_allow_html=True)

    dokumen = st.file_uploader(
        "Upload Dokumen",
        type=["pdf", "txt", "csv", "xlsx", "xls"]
    )

# ==============================
# TOMBOL PROSES
# ==============================
proses = st.button("🔍 PROSES EKSTRAKSI", use_container_width=True)

if proses:
    progress_bar = st.progress(0)
    status_text = st.empty()

    status_text.text("Memulai proses ekstraksi... 10%")
    progress_bar.progress(10)

    status_text.text("Membaca input tanaman... 30%")
    progress_bar.progress(30)

    status_text.text("Membaca dokumen... 50%")
    progress_bar.progress(50)

    status_text.text("Mencocokkan entitas bioaktif... 75%")
    progress_bar.progress(75)

    status_text.text("Membangun hasil ekstraksi... 100%")
    progress_bar.progress(100)

    st.success("Proses ekstraksi selesai.")

    # ==============================
    # HASIL EKSTRAKSI SEMENTARA
    # ==============================
    st.markdown("""
    <div class="result-card">
        <h2>📋 Hasil Ekstraksi Informasi Bioaktif</h2>
    </div>
    """, unsafe_allow_html=True)

    st.markdown('<div class="card-lilac">', unsafe_allow_html=True)

    st.write("**Input Tanaman / Kalimat:**")
    st.write(teks if teks else "Belum ada input teks.")

    st.write("**Dokumen yang diupload:**")
    if dokumen:
        st.write(dokumen.name)
    else:
        st.write("Belum ada dokumen yang diupload.")

    st.write("**Nama Tanaman:** Belum dikoneksikan ke dataset")
    st.write("**Nama Latin:** Belum dikoneksikan ke dataset")
    st.write("**Zat Bioaktif:** Belum dikoneksikan ke dataset")
    st.write("**Khasiat / Efek Terapeutik:** Belum dikoneksikan ke dataset")

    st.markdown('</div>', unsafe_allow_html=True)

    # ==============================
    # RELATION EXTRACTION
    # ==============================
    st.markdown("""
    <div class="result-card">
        <h2>🔗 Bioactive Relation Extraction</h2>
        <p><b>Tanaman</b> → contains_bioactive_compound → <b>Senyawa Bioaktif</b></p>
        <p><b>Tanaman</b> → has_therapeutic_effect → <b>Khasiat</b></p>
        <p><b>Tanaman</b> → uses_part → <b>Bagian Tanaman</b></p>
    </div>
    """, unsafe_allow_html=True)

    # ==============================
    # HERBKG
    # ==============================
    st.markdown("""
    <div class="result-card">
        <h2>🕸️ Enhanced Herb Knowledge Graph 2.0</h2>
        <p>
        HerbKG 2.0 akan menampilkan hubungan antara tanaman, nama latin,
        bagian tanaman, senyawa bioaktif, khasiat, dosis, cara pengolahan,
        dan sumber data.
        </p>
    </div>
    """, unsafe_allow_html=True)

else:
    st.info("Masukkan nama tanaman atau upload dokumen, lalu klik tombol PROSES EKSTRAKSI.")
