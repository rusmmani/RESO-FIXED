RESO HAIRCUT — SPREADSHEET INTEGRATION

Perubahan utama:
1. index.html -> POST /api/reso -> Google Apps Script -> Spreadsheet.
2. admin.html -> GET/POST /api/reso -> data booking yang sama.
3. Data lama di tab `Bookings` tetap dibaca.
4. Admin sekarang punya tombol Edit (nama, WhatsApp, catatan, status) selain reschedule/selesai/hapus.
5. API Vercel memakai URL Apps Script /exec sebagai fallback, jadi tidak wajib env GAS_URL.

PENTING — DEPLOY APPS SCRIPT
URL yang dikirim adalah /dev. Untuk website produksi gunakan Web App URL /exec.

Di Apps Script:
- Deploy > New deployment
- Type: Web app
- Execute as: Me
- Who has access: Anyone
- Deploy
- Copy URL yang berakhiran /exec.

Jika URL /exec berbeda dari yang tertanam di api/reso.js, set Vercel Environment Variable:
GAS_URL=https://script.google.com/macros/s/DEPLOYMENT_ID/exec
lalu redeploy Vercel.

SPREADSHEET
Spreadsheet ID sudah tertanam di Code.gs:
11m-xtfldUCVuJQsB9-FUCGMK8V4d5QpekZPOndANWOQ

Tab yang digunakan:
- Bookings (legacy/data lama)
- Januari 2026, Februari 2026, dst. (data bulanan)
- Capsters
- Promos

Catatan:
- Sistem otomatis membuat tab bulan aktif.
- Data lama di Bookings tetap muncul di admin.
- Saat data lama diedit/reschedule, sistem dapat memindahkannya ke tab bulan yang sesuai.


UPDATE TERBARU
- index.html sekarang menggunakan foto capster baru dari folder assets/REZKY.png dan assets/IQBAL.png.
- Dashboard admin memiliki tombol Capster di topbar. Klik untuk membuka popup Capster.
- Popup Capster menampilkan foto, status ON/OFF, dan toggle yang langsung menyimpan status ke sheet Capsters.
- Tombol + Booking Baru sekarang membuka form Booking Manual langsung di dashboard.
- Booking manual memakai action createBooking yang sudah ada di Code.gs, sehingga tetap tersimpan ke tab bulan Google Spreadsheet dan mengikuti validasi slot/capster aktif.


UPDATE 2026-09-10
- Logo RESO sekarang tertanam langsung sebagai Base64 di index.html dan admin.html, jadi tidak bergantung pada file storage/assets untuk logo.
- Halaman utama menghitung slot terdekat secara live berdasarkan jam buka 10:00–22:00; untuk hari ini, slot dimulai dari jam penuh berikutnya (contoh 12:xx -> 13.00).
- Dashboard menampilkan cache lokal seketika, lalu sinkron otomatis ke Spreadsheet.
- Dashboard memiliki tombol Refresh manual dan polling otomatis.
- Apps Script menambahkan endpoint getDashboardData untuk mengambil booking + settings + daftar bulan dalam satu request, dengan cache singkat agar lebih cepat.
