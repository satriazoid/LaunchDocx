import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import re

# --- KONFIGURASI TEMA SOFT ---
st.set_page_config(page_title="White Paper Generator", page_icon="📝")

st.markdown("""
    <style>
    .main { background-color: #fdfbfb; }
    .stButton>button {
        background-color: #a8dadc;
        color: #1d3557;
        border-radius: 10px;
        border: none;
        padding: 10px 24px;
        font-weight: bold;
    }
    <div class="hero-box">
    <h1>White Paper Automation</h1>
    <p>Buat dokumen profesional dengan satu klik</p>
    </div>
    .stTextInput>div>div>input { background-color: #f1faee; }
    h1 { color: #457b9d; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
    </style>
    """, unsafe_allow_html=True)

# --- FUNGSI MEMBERSIHKAN TEKS DARI AI ---
def clean_ai_text(text):
    if not text:
        return ""
    # 1. Menghapus simbol markdown seperti asteris (**) untuk bold dari AI
    text = text.replace("**", "").replace("__", "")
    # 2. Menghapus spasi berlebih di awal/akhir baris
    lines = [line.strip() for line in text.split('\n')]
    # 3. Menggabungkan kembali dengan satu spasi (menghindari baris baru yang tidak perlu)
    cleaned_text = " ".join(lines)
    # 4. Menghapus double space yang sering muncul
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text)
    return cleaned_text

# --- FUNGSI GARIS FULL WIDTH ---
def add_full_line(paragraph):
    p_element = paragraph._element
    p_pr = p_element.get_or_add_pPr()
    p_border = OxmlElement('w:pBdr')
    bottom_border = OxmlElement('w:bottom')
    bottom_border.set(qn('w:val'), 'single')
    bottom_border.set(qn('w:sz'), '6')
    bottom_border.set(qn('w:space'), '1')
    bottom_border.set(qn('w:color'), '000000')
    p_border.append(bottom_border)
    p_pr.append(p_border)

# --- FUNGSI GENERATE DOCX ---
def generate_docx(judul, penulis, konten_dict):
    doc = Document()
    
    def set_font(run, size, bold=False):
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        run.font.size = Pt(size)
        run.font.bold = bold
        run.font.color.rgb = RGBColor(0, 0, 0)

    # 1. White Paper (Size 16, Bold, Center)
    wp_p = doc.add_paragraph()
    wp_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    wp_run = wp_p.add_run('White Paper')
    set_font(wp_run, 16, True)

    # 2. Judul (Size 14, No Bold, Center)
    t_p = doc.add_paragraph()
    t_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    t_run = t_p.add_run(judul.upper())
    set_font(t_run, 14, False)

    # 3. Penulis
    p_p = doc.add_paragraph()
    p_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_run = p_p.add_run(f"Penulis: {penulis}")
    set_font(p_run, 12, False)

    sections_order = ["Eksekutif", "Pendahuluan", "Pernyataan Masalah", "Metodologi", "Solusi", "Manfaat dan Hasil", "Kesimpulan", "Referensi dan Lampiran"]

    for heading in sections_order:
        # Garis Full Width
        line_para = doc.add_paragraph()
        add_full_line(line_para)
        
        # Heading (Bold, Left)
        h2_p = doc.add_paragraph()
        h2_run = h2_p.add_run(heading)
        set_font(h2_run, 12, True)
        
        # ISI PARAGRAF (Rata Kanan Kiri / Justify)
        raw_text = konten_dict.get(heading, "")
        cleaned_text = clean_ai_text(raw_text) # Membersihkan teks AI
        
        body_p = doc.add_paragraph()
        body_p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY # <--- RATA KANAN KIRI
        body_run = body_p.add_run(cleaned_text if cleaned_text else "Teks belum diisi...")
        set_font(body_run, 12, False)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- UI (Sesuai Struktur Anda) ---
st.title("📝 White Paper Automation")
st.subheader("Buat dokumen profesional dengan satu klik")

with st.container():
    col1, col2 = st.columns(2)
    with col1:
        judul_input = st.text_input("Judul White Paper", placeholder="Contoh: Masa Depan AI di Indonesia")
    with col2:
        penulis_input = st.text_input("Nama Penulis", placeholder="Nama Lengkap Anda")

    st.markdown("---")
    
    sections = ["Eksekutif", "Pendahuluan", "Pernyataan Masalah", "Metodologi", "Solusi", "Manfaat dan Hasil", "Kesimpulan", "Referensi dan Lampiran"]
    
    konten_user = {}
    for sec in sections:
        konten_user[sec] = st.text_area(f"Isi Bagian {sec}", placeholder=f"Tulis atau tempel teks dari AI di sini...")

    if st.button("Generate & Download Document"):
        if judul_input and penulis_input:
            doc_file = generate_docx(judul_input, penulis_input, konten_user)
            st.success("Dokumen berhasil dibuat & dibersihkan!")
            st.download_button(
                label="⬇️ Klik untuk Mengunduh (.docx)",
                data=doc_file,
                file_name=f"White_Paper_{judul_input.replace(' ', '_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
        else:
            st.error("Mohon isi Judul dan Nama Penulis.")