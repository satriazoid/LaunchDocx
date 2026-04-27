import streamlit as st
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from io import BytesIO
import re

# --- KONFIGURASI TEMA UI ---
st.set_page_config(page_title="UNPAM Doc Generator", page_icon="🎓")

# --- FUNGSI MEMBERSIHKAN TEKS ---
def clean_ai_text(text):
    if not text: return ""
    text = text.replace("**", "").replace("__", "")
    return text.strip()

# --- FUNGSI GENERATE DOCX ---
def generate_unpam_docx(judul, penulis, nim, konten_dict):
    doc = Document()
    
    # Fungsi Helper Font
    def format_run(run, size, bold=False):
        run.font.name = 'Times New Roman'
        run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Times New Roman')
        run.font.size = Pt(size)
        run.font.bold = bold

    # --- 1. HALAMAN COVER ---
    cp = doc.add_paragraph()
    cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_run(cp.add_run(judul.upper() + "\n\n\n\n\n\n\n"), 14, True)
    
    cp_penulis = doc.add_paragraph()
    cp_penulis.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_run(cp_penulis.add_run(f"Oleh:\n{penulis.upper()}\nNIM: {nim}\n\n\n\n\n\n"), 12, True)
    
    cp_univ = doc.add_paragraph()
    cp_univ.alignment = WD_ALIGN_PARAGRAPH.CENTER
    format_run(cp_univ.add_run("PROGRAM STUDI TEKNIK INFORMATIKA\nUNIVERSITAS PAMULANG\n2026"), 14, True)
    doc.add_page_break()

    # --- 2. HALAMAN FORMAL (PENGANTAR, DAFTAR ISI, DLL) ---
    halaman_formal = ["KATA PENGANTAR", "DAFTAR ISI", "DAFTAR TABEL", "DAFTAR GAMBAR"]
    for hal in halaman_formal:
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_run(p.add_run(hal), 14, True)
        
        content = doc.add_paragraph()
        content.paragraph_format.line_spacing = 1.5
        content.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        format_run(content.add_run(clean_ai_text(konten_dict.get(hal, "..."))), 12, False)
        doc.add_page_break()

    # --- 3. BAB 1 - 5 ---
    bab_list = [
        ("BAB I", "PENDAHULUAN"),
        ("BAB II", "LANDASAN TEORI"),
        ("BAB III", "METODOLOGI PENELITIAN"),
        ("BAB IV", "HASIL DAN PEMBAHASAN"),
        ("BAB V", "KESIMPULAN DAN SARAN")
    ]

    for code, nama in bab_list:
        # Judul Bab
        p_bab = doc.add_paragraph()
        p_bab.alignment = WD_ALIGN_PARAGRAPH.CENTER
        format_run(p_bab.add_run(f"{code}\n{nama}"), 14, True)
        
        # Isi Bab
        p_isi = doc.add_paragraph()
        p_isi.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p_isi.paragraph_format.line_spacing = 1.5
        # Mengambil isi berdasarkan nama Bab
        raw_text = konten_dict.get(nama, "Tulis isi bab di sini...")
        format_run(p_isi.add_run(clean_ai_text(raw_text)), 12, False)
        
        doc.add_page_break()

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer

# --- UI STREAMLIT ---
st.title("🎓 UNPAM Document Generator")
st.write("Format otomatis: Times New Roman, 1.5 Spacing, Judul Bab 14pt.")

with st.expander("Informasi Dokumen", expanded=True):
    col1, col2, col3 = st.columns(3)
    judul = col1.text_input("Judul Tugas Akhir")
    nama = col2.text_input("Nama Lengkap")
    nim = col3.text_input("NIM")

st.markdown("### Isi Konten Halaman")
tabs = st.tabs(["Formal", "Bab 1-3", "Bab 4-5"])

konten_user = {}

with tabs[0]:
    konten_user["KATA PENGANTAR"] = st.text_area("Isi Kata Pengantar")
    konten_user["DAFTAR ISI"] = st.text_area("Daftar Isi (Manual/Placeholder)")
    konten_user["DAFTAR TABEL"] = st.text_area("Daftar Tabel")
    konten_user["DAFTAR GAMBAR"] = st.text_area("Daftar Gambar")

with tabs[1]:
    konten_user["PENDAHULUAN"] = st.text_area("Isi BAB I")
    konten_user["LANDASAN TEORI"] = st.text_area("Isi BAB II")
    konten_user["METODOLOGI PENELITIAN"] = st.text_area("Isi BAB III")

with tabs[2]:
    konten_user["HASIL DAN PEMBAHASAN"] = st.text_area("Isi BAB IV")
    konten_user["KESIMPULAN DAN SARAN"] = st.text_area("Isi BAB V")

if st.button("Generate Dokumen Formal"):
    if judul and nama:
        file_docx = generate_unpam_docx(judul, nama, nim, konten_user)
        st.success("Dokumen siap diunduh!")
        st.download_button(
            label="⬇️ Download File .docx",
            data=file_docx,
            file_name=f"Tugas_Akhir_{nama}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Judul dan Nama tidak boleh kosong.")