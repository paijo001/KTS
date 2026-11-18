# Panduan Proyek KTS

Gunakan file ini sebagai ringkasan saat melanjutkan pekerjaan. Baca file yang berkaitan langsung dengan tugas; jangan membaca ulang seluruh proyek tanpa kebutuhan.

## Struktur

- `index.php`: aplikasi PHP satu halaman; impor spreadsheet, pembuatan tabel, pencarian/filter/paginasi pelanggan, status pembayaran, dan tautan invoice WhatsApp.
- `style.css`: seluruh styling antarmuka, termasuk layout responsif smartphone.
- `config.php`: koneksi MySQL dan konfigurasi aplikasi yang sudah ada. Jangan menaruh atau menyalin kredensial ke dokumentasi.
- `PREV AGUSTUS NO HP.xlsx`: contoh file master pelanggan.

## Database dan alur data

- Aplikasi memakai koneksi `$conn` dari `config.php` (database `data_sblw`).
- `index.php` membuat tabel `billing_customers` dan `billing_app_meta` jika belum ada; MySQL harus aktif dan database sudah tersedia.
- File master bulanan harus `.xlsx`, dengan header minimal `IDPEL` dan `NAMA`. Header yang digunakan bila tersedia: `KOKED`, `ALAMAT`, `TARIF`, `DAYA`, `RPTAG`, `RPBK`, `TOTAL`, `KONTAKPELANGGAN`, `NO WA`.
- Impor master mengganti seluruh data pelanggan dalam transaksi. Status direset menjadi `unknown`; jika impor gagal, transaksi di-rollback agar data lama tetap aman.
- Impor harian menerima `.xlsx` atau `.csv`. File dapat memiliki header `IDPEL` atau berupa daftar IDPEL tanpa header. IDPEL yang ada pada file ditandai `unpaid`; pelanggan master yang tidak ada pada file ditandai `paid`.
- Status awal `unknown` berarti belum ada file harian sejak master terakhir diunggah. Tombol invoice hanya muncul untuk pelanggan `unpaid`.
- KODED lengkap ditampilkan di samping IDPEL. Pelanggan dikelompokkan berdasarkan 6 karakter pertama kolom KOKED; pilihan sorting mengurutkan data di dalam tiap grup.
- Pencarian mendukung IDPEL, KODED, nama, alamat, dan nomor WhatsApp; sorting tersedia untuk KODED, nama, IDPEL, dan jumlah tagihan.
- Tombol WhatsApp per pelanggan membuka `wa.me` dengan pesan invoice yang menyapa Bpk/Ibu dan mencantumkan batas pembayaran tanggal 20 setiap bulan; pengguna tetap meninjau dan mengirim pesan dari WhatsApp.
- Saat tombol WA dibuka untuk pelanggan belum lunas, klik disimpan pada tabel `billing_whatsapp_clicks`; tombol berubah warna selama tanggal klik sama dengan hari ini, lalu otomatis kembali hijau pada tanggal berikutnya.
- Penanda klik WA hanya menandakan tautan dibuka, bukan bukti pesan sudah terkirim.
- Tombol tema di header berganti antara terang/gelap; pilihan disimpan di `localStorage` browser dengan kunci `kts-theme`.
- URL stylesheet memakai `filemtime` sebagai versi cache; unggah `index.php` dan `style.css` bersama-sama saat deployment.
- Pencarian/filter dilakukan dari server. Jumlah data per halaman: 25.

## Perubahan dan verifikasi

- Jaga perubahan tetap terfokus; styling baru ditempatkan di `style.css`, bukan inline di halaman.
- Jangan mengubah semantik status pelunasan tanpa memastikan bahwa file harian tetap merupakan daftar pelanggan yang belum lunas.
- Validasi sintaks PHP dengan `php -l index.php`.
- Parser dapat diperiksa dengan menjalankan `readXlsxRows('PREV AGUSTUS NO HP.xlsx')` dan `readCsvRows(...)` dari konteks PHP setelah bootstrap aplikasi.
- Tampilan memerlukan MySQL untuk berfungsi penuh. Jika koneksi belum tersedia, halaman menampilkan pesan bahwa layanan MySQL/database perlu dijalankan atau dikonfigurasi.
- Pada pemeriksaan terakhir, PHP CLI dan ekstensi `mysqli`, `zip`, serta `dom` tersedia; layanan MySQL belum aktif.
