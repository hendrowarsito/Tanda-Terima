# Tanda Terima (prototipe)

Pengganti OmniForm: penomoran otomatis, master klien/pengantar, item dinamis, PDF tanda terima + amplop,
riwayat, status pengiriman, dasbor "belum kembali", cadangan JSON.
Penyimpanan: **Google Sheets** (utama) atau SQLite (cadangan, hilang saat restart di Streamlit Cloud).

## Tampilan & tema (baru)
- **Navigasi atas** (Buat · Riwayat · Klien & Pengantar · Pengaturan) dengan logo; di layar sempit/ponsel menu masuk ke panel samping.
- **Buat:** formulir bernomor (Pengiriman → Penerima → Daftar dokumen) di kiri, **pratinjau A4 yang menempel** di kanan.
- **Riwayat:** kartu ringkasan (total, bulan ini, menunggu diterima, lewat 7 hari), tabel terurut dari yang terbaru dengan label status berwarna;
  **klik satu baris** untuk membuka detail, mengubah status, atau mencetak.
- **Klien & Pengantar** dan **Pengaturan** memakai tab; pesan "tersimpan" tampil sebagai notifikasi.
- Halaman masuk (bila `app_password` diisi) berlogo dan mendukung tombol Enter.
- Warna utama mengikuti logo (`#2E1F6E`). Tema diatur di `.streamlit/config.toml` (**wajib ikut diunggah ke GitHub**) dan CSS di dalam `app.py`.
  Tema dikunci mode terang agar warna konsisten.
- File font `TCCEB.TTF` tetap diperlukan (kop surat PDF dan tulisan "Tanda Terima" di bar atas).

**File yang diunggah ke repo:** `app.py`, `requirements.txt`, `TCCEB.TTF`, `.streamlit/config.toml` (opsional `logo.png`).

## 1. Siapkan Google Sheets (sekali saja)
1. Buka https://console.cloud.google.com → buat project baru.
2. *APIs & Services → Library*: aktifkan **Google Sheets API** dan **Google Drive API**.
3. *APIs & Services → Credentials → Create credentials → Service account* → beri nama → selesai.
4. Buka service account itu → tab **Keys → Add key → JSON**. File JSON terunduh (rahasia, jangan di-commit ke GitHub).
5. Buat **Google Sheet kosong** di akun Google Anda. Klik **Share**, tambahkan alamat `client_email`
   dari file JSON, peran **Editor**. Jangan bagikan ke orang lain / "anyone with the link".
6. Salin ID sheet dari URL: `https://docs.google.com/spreadsheets/d/`**`ID_INI`**`/edit`.

## 2. Deploy ke Streamlit Community Cloud
1. Unggah `app.py`, `requirements.txt`, `README.md` ke repo GitHub.
2. share.streamlit.io → New app → pilih repo → main file `app.py`.
3. *Settings → Secrets* (salin nilai dari file JSON):
```toml
app_password = "kata-sandi-pilihan-anda"
gsheet_id = "ID_SHEET_ANDA"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "nama@project.iam.gserviceaccount.com"
client_id = "..."
token_uri = "https://oauth2.googleapis.com/token"
```
Pada `private_key`, biarkan `\n` apa adanya (jangan diganti baris baru) dan tetap diapit tanda kutip.

Saat pertama dibuka, aplikasi membuat empat tab sendiri di Sheet: `clients`, `couriers`, `receipts`, `settings`.
Jangan ubah nama tab atau urutan/judul kolom. Mengedit isi sel manual boleh, tapi kolom `id` jangan diubah.

## Pratinjau & cetak (kertas A4)
- Satu lembar **A4**: **setengah atas = tanda terima**, **setengah bawah = amplop** berbingkai (garis putus-putus = tempat memotong).
  Pratinjau di aplikasi menampilkan lembar A4 utuh, persis seperti yang tercetak.
- **Buat:** pratinjau ikut berubah saat Anda mengisi; nomor final ditetapkan saat **Simpan**, lalu muncul blok cetak.
  **Riwayat:** pilih tanda terima untuk pratinjau dan cetak ulang.
- Pilih yang dicetak: **Tanda terima + amplop**, **Tanda terima saja**, atau **Amplop saja**. Posisinya tetap, sehingga separuh
  kertas yang masih kosong bisa dipakai lagi untuk cetakan berikutnya.
- Bila daftar dokumen panjang, tanda terima berlanjut ke lembar berikutnya (amplop hanya di lembar pertama).
- **🖨️ Cetak** membuka jendela baru berisi halaman siap cetak lalu memanggil dialog cetak (izinkan pop-up). Jika dialog tidak muncul,
  tekan tombol Cetak di jendela itu. Di dialog: ukuran **A4**, skala **100%**, matikan *header & footer*.
- Cadangan: **📄 Buka PDF** lalu Cmd/Ctrl+P, atau **Unduh PDF**.

## Status "Diterima"
Di **Riwayat**, memilih status **Diterima** memunculkan isian **Penerima** (wajib), **Tanggal**, dan **Jam diterima**.
Datanya tersimpan, tampil di tabel riwayat, dan tercetak di bawah kolom *Diterima oleh*.
Database lama (SQLite maupun Google Sheets) otomatis ditambah tiga kolom baru saat aplikasi dibuka; data lama tidak berubah.

## Logo
Logo KJPP sudah tertanam di `app.py`. Untuk menggantinya, simpan file `logo.png` di folder yang sama dengan `app.py`.

## Batasan
- Kuota Google API ±60 pembacaan/menit per pengguna; aplikasi men-cache data 20 detik, cukup untuk beberapa pengguna.
- Bila banyak orang menyimpan bersamaan, nomor tetap dicek ulang agar tidak ganda, tapi untuk beban tinggi
  gunakan database sungguhan (Postgres).

## Jalankan lokal
`pip install -r requirements.txt && streamlit run app.py` (tanpa Secrets → SQLite lokal).
