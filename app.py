import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from io import BytesIO
import re

# --- FUNGSI CLEANSING ---
def clean_ai_content(text):
    if not text: return ""
    text = text.replace("**", "").replace("__", "").replace("*", "").replace("_", "")
    text = re.sub(r'#+\s?', '', text)
    text = text.replace("--", "").replace("—", "").replace("–", "")
    lines = text.split('\n')
    cleaned_lines = [re.sub(r'^[•\-\s]+', '', line) for line in lines]
    return '\n'.join(cleaned_lines).strip()

# --- FUNGSI STYLING HEADING & PARAGRAF ---
def set_heading_style(paragraph, size, bold=True):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER if size == 14 else WD_ALIGN_PARAGRAPH.LEFT
    paragraph.paragraph_format.line_spacing = 1.5
    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    run.font.name = 'Times New Roman'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = RGBColor(0, 0, 0) # Wajib Hitam

def add_body_text(doc, text):
    cleaned = clean_ai_content(text)
    if not cleaned: return
    p = doc.add_paragraph(cleaned)
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing = 1.5
    for run in p.runs:
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(0, 0, 0)

# --- CORE GENERATOR ---
def generate_formal_document(data):
    doc = Document()
    
    # 1. COVER (Non-Heading)
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_judul = cp.add_run(data['judul'].upper() + "\n\n\n\n\n\n")
    set_heading_style(cp, 14, True)
    
    cp2 = doc.add_paragraph()
    cp2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_mhs = cp2.add_run(f"Oleh:\n{data['nama'].upper()}\nNIM: {data['nim']}\n\n\n\n\n\n")
    set_heading_style(cp2, 12, True)
    
    cp3 = doc.add_paragraph()
    cp3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_univ = cp3.add_run("PROGRAM STUDI TEKNIK INFORMATIKA\nUNIVERSITAS PAMULANG\n2026")
    set_heading_style(cp3, 14, True)
    doc.add_page_break()

    # 2. HALAMAN FORMAL (HEADING 1)
    formal_pages = ["KATA PENGANTAR", "DAFTAR ISI", "DAFTAR TABEL", "DAFTAR GAMBAR"]
    for title in formal_pages:
        h = doc.add_heading(title, level=1)
        set_heading_style(h, 14, True)
        if data.get(title):
            add_body_text(doc, data[title])
        doc.add_page_break()

    # 3. BAB I - V & SUB-BAB (HEADING 1 & 2)
    struktur = [
        ("BAB I", "PENDAHULUAN", ["1.1 Latar Belakang", "1.2 Rumusan Masalah", "1.3 Tujuan Penelitian"]),
        ("BAB II", "LANDASAN TEORI", ["2.1 Teori Umum", "2.2 Teori Khusus"]),
        ("BAB III", "METODOLOGI PENELITIAN", ["3.1 Alur Penelitian", "3.2 Teknik Data"]),
        ("BAB IV", "HASIL DAN PEMBAHASAN", ["4.1 Implementasi", "4.2 Pembahasan"]),
        ("BAB V", "PENUTUP", ["5.1 Kesimpulan", "5.2 Saran"])
    ]

    for code, name, subs in struktur:
        h_bab = doc.add_heading(f"{code} {name}", level=1)
        set_heading_style(h_bab, 14, True)
        
        for sub in subs:
            h_sub = doc.add_heading(sub, level=2)
            set_heading_style(h_sub, 12, True)
            add_body_text(doc, data.get(sub, ""))
        doc.add_page_break()

    # 4. DAFTAR PUSTAKA (HEADING 1)
    h_ref = doc.add_heading("DAFTAR PUSTAKA", level=1)
    set_heading_style(h_ref, 14, True)
    add_body_text(doc, data.get("DAFTAR PUSTAKA", ""))

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- STREAMLIT UI ---
st.title("🎓 UNPAM Doc Engine (Heading Support)")

with st.sidebar:
    st.header("Konfigurasi")
    judul = st.text_input("Judul", "Analisis Strategi...")
    nama = st.text_input("Nama")
    nim = st.text_input("NIM")

tabs = st.tabs(["Formal", "Bab I-II", "Bab III-IV", "Bab V & Pustaka"])
konten = {'judul': judul, 'nama': nama, 'nim': nim}

with tabs[0]:
    konten["KATA PENGANTAR"] = st.text_area("Kata Pengantar")
    konten["DAFTAR ISI"] = st.text_area("Daftar Isi (Opsional)")
    
with tabs[1]:
    konten["1.1 Latar Belakang"] = st.text_area("1.1 Latar Belakang")
    konten["1.2 Rumusan Masalah"] = st.text_area("1.2 Rumusan Masalah")
    konten["1.3 Tujuan Penelitian"] = st.text_area("1.3 Tujuan Penelitian")
    konten["2.1 Teori Umum"] = st.text_area("2.1 Teori Umum")
    konten["2.2 Teori Khusus"] = st.text_area("2.2 Teori Khusus")

with tabs[2]:
    konten["3.1 Alur Penelitian"] = st.text_area("3.1 Alur Penelitian")
    konten["3.2 Teknik Data"] = st.text_area("3.2 Teknik Data")
    konten["4.1 Implementasi"] = st.text_area("4.1 Implementasi")
    konten["4.2 Pembahasan"] = st.text_area("4.2 Pembahasan")

with tabs[3]:
    konten["5.1 Kesimpulan"] = st.text_area("5.1 Kesimpulan")
    konten["5.2 Saran"] = st.text_area("5.2 Saran")
    konten["DAFTAR PUSTAKA"] = st.text_area("Daftar Pustaka")

if st.button("Generate Formal Document", use_container_width=True):
    if not nama or not nim:
        st.warning("Lengkapi Nama dan NIM di sidebar.")
    else:
        file = generate_formal_document(konten)
        st.success("Berhasil! Heading sudah aktif di Navigation Pane Word.")
        st.download_button("Download .docx", file, f"Laporan_{nama}.docx")