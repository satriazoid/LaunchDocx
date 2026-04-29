# Universal Academic Document Engine

Universal Academic Document Engine adalah aplikasi berbasis web yang dibangun dengan Streamlit untuk membantu pelajar, mahasiswa, dan profesional dalam menyusun dokumen formal. Engine ini secara otomatis mengatur format tipografi, struktur heading, dan pembersihan artefak teks yang dihasilkan oleh AI (seperti sisa-sisa markdown dari ChatGPT atau Gemini).

## Fitur Utama

- **Modular Chapter System**: Pengguna dapat menentukan sendiri jumlah bab (1 hingga 7 bab) sesuai dengan kebutuhan laporan atau makalah.
- **Dynamic Heading Detection**: Sistem secara otomatis mendeteksi penomoran seperti 1.1, 2.1.1, atau 3.2 sebagai Heading resmi Microsoft Word yang terintegrasi dengan Navigation Pane.
- **Auto Table of Contents**: Menyisipkan field Daftar Isi otomatis yang dapat diperbarui (update field) langsung di dalam Microsoft Word.
- **Deep Cleansing AI**: Membersihkan simbol markdown (**), sisa heading (#), serta karakter pemisah (--) secara otomatis sebelum dokumen dicetak.
- **Standard Academic Formatting**:
  1. Font: Times New Roman.
  2. Ukuran: 12pt (Isi) dan 14pt (Judul).
  3. Spasi: 1.5 Line Spacing.
  4. Perataan: Justify (Rata Kanan-Kiri).
  5. Alinea: First Line Indent 1cm.
- **Multi-Author Support**: Mendukung penulisan nama anggota kelompok secara vertikal yang rapi pada halaman cover.
- **Modular Pages**: Tersedia opsi On/Off untuk Kata Pengantar, Daftar Isi, Daftar Tabel, Daftar Gambar, dan Daftar Pustaka.

## Persyaratan Sistem

Pastikan Anda telah menginstal Python 3.8 atau versi yang lebih baru di sistem Anda.

## Instalasi

1. Clone repository ini:
   ```bash
   git clone [https://github.com/username/universal-doc-engine.git](https://github.com/username/universal-doc-engine.git)
   cd universal-doc-engine