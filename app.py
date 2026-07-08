import streamlit as st

st.set_page_config(
    page_title="HyTBIONEX",
    page_icon="🌿",
    layout="wide"
)

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
    font-size: 52px;
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

.result-card {
    background: #f3e8ff;
    padding: 22px;
    border-radius: 20px;
    border: 2px solid #c084fc;
    color: #111111;
    margin-top: 20px;
    box-shadow: 0 8px 18px rgba(0,0,0,0.12);
}
</style>
""", unsafe_allow_html=True)

# SIDEBAR
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

# HEADER
st.markdown("""
<div class="main-title">
    <h1>🌿 HyTBIONEX</h1>
    <h3>Hybrid Transformer for Bioactive Information Extraction</h3>
    <p>Analisis Bioaktif Tanaman Herbal Indonesia & Enhanced Herb Knowledge Graph 2.0</p>
</div>
""", unsafe_allow_html=True)

# WELCOME
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

# INPUT
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

# PROSES
if st.button("🔍 PROSES EKSTRAKSI", use_container_width=True):
    progress = st.progress(0)
    status = st.empty()

    status.text("Memulai proses ekstraksi... 10%")
    progress.progress(10)

    status.text("Membaca input tanaman... 30%")
    progress.progress(30)

    status.text("Membaca dokumen... 50%")
    progress.progress(50)

    status.text("Mencocokkan entitas bioaktif... 75%")
    progress.progress(75)

    status.text("Membangun hasil ekstraksi... 100%")
    progress.progress(100)

    st.success("Proses ekstraksi selesai.")

    st.markdown("""
    <div class="result-card">
        <h2>📋 Hasil Ekstraksi Informasi Bioaktif</h2>
    </div>
    """, unsafe_allow_html=True)

    st.write("**Input Tanaman / Kalimat:**")
    st.write(teks if teks else "Belum ada input teks.")

    st.write("**Dokumen yang diupload:**")
    if dokumen:
        st.write(dokumen.name)
    else:
        st.write("Belum ada dokumen yang diupload.")

    st.markdown("""
    <div class="result-card">
        <h3>🌿 Entitas Bioaktif</h3>
        <p><b>Nama Tanaman:</b> Belum dikoneksikan ke dataset</p>
        <p><b>Nama Latin:</b> Belum dikoneksikan ke dataset</p>
        <p><b>Zat Bioaktif:</b> Belum dikoneksikan ke dataset</p>
        <p><b>Khasiat / Efek Terapeutik:</b> Belum dikoneksikan ke dataset</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="result-card">
        <h3>🔗 Bioactive Relation Extraction</h3>
        <p><b>Tanaman</b> → contains_bioactive_compound → <b>Senyawa Bioaktif</b></p>
        <p><b>Tanaman</b> → has_therapeutic_effect → <b>Khasiat</b></p>
        <p><b>Tanaman</b> → uses_part → <b>Bagian Tanaman</b></p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("""
    <div class="result-card">
        <h3>🕸️ Enhanced Herb Knowledge Graph 2.0</h3>
        <p>
        HerbKG 2.0 akan menampilkan hubungan antara tanaman, nama latin,
        bagian tanaman, senyawa bioaktif, khasiat, dosis, cara pengolahan,
        dan sumber data.
        </p>
    </div>
    """, unsafe_allow_html=True)

else:
    st.info("Masukkan nama tanaman atau upload dokumen, lalu klik tombol PROSES EKSTRAKSI.")
