SC MONITORING DASHBOARD — JAKARTANOTEBOOK THEME

MENJALANKAN APLIKASI
1. Ekstrak ZIP ke folder permanen.
2. Buka PowerShell pada folder tersebut.
3. Buat virtual environment (sekali saja):
   py -m venv .venv
4. Install kebutuhan:
   .\.venv\Scripts\python.exe -m pip install -r requirements.txt
5. Jalankan:
   .\.venv\Scripts\python.exe -m streamlit run app.py

KONEKSI GOOGLE SHEETS
1. Pilih sumber data "Google Sheets (otomatis)" di sidebar.
2. Pastikan spreadsheet Google Sheets dapat dibaca oleh akun/aplikasi:
   Share > General access > Anyone with the link > Viewer.
3. Salin URL standar Google Sheets (docs.google.com/spreadsheets/d/...).
4. Tempel URL pada kotak URL Google Sheets.
5. Aplikasi mengunduh seluruh workbook sebagai XLSX. Semua tab Master 2025/2026/2027
   dan Case Pic perlu tetap ada dengan nama sheet dan header kolom yang sama.
6. Refresh otomatis berjalan setiap 5 menit jika diaktifkan. Tombol "Refresh data sekarang"
   memaksa aplikasi mengambil data baru.
7. Untuk spreadsheet privat/terbatas organisasi, jangan membuka akses publik. Integrasi
   perlu autentikasi Google/service account dan izin akses yang sesuai.

SUMBER DATA CADANGAN
- Jika URL Google Sheets belum diisi, aplikasi menggunakan file Excel bawaan.
- Pilihan "Upload Excel (manual)" tetap tersedia.

PERUBAHAN PENGEMBANGAN
- Tema visual navy-biru-oranye ala JakartaNotebook / operasi gudang.
- Menu Kualitas Data Case: case tanpa Tanggal Complaint, tahun sebelumnya/setelahnya,
  dan klasifikasi kosong/di luar dua kategori.
- Rekap Case Tahunan menggunakan Tanggal Complaint.
- Monitoring Case tetap menggunakan Tanggal Diterima.
- File Excel tetap disertakan sebagai data awal/cadangan.
