# Tanda Terima (prototipe)

Pengganti OmniForm: penomoran otomatis, master klien/pengantar, item dinamis, PDF tanda terima + amplop,
riwayat, status pengiriman, dasbor "belum kembali", cadangan JSON.
Penyimpanan: **Google Sheets** (utama) atau SQLite (cadangan, hilang saat restart di Streamlit Cloud).

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

## Pratinjau & cetak
- **Buat:** pratinjau tampil langsung di bawah formulir dan ikut berubah saat Anda mengisi. Nomor final baru
  ditetapkan saat **Simpan**; setelah itu muncul blok cetak.
- **Riwayat:** pilih tanda terima, lalu pratinjau dan tombol cetak muncul di bawahnya.
- Pilih yang dicetak: **Tanda terima / Amplop / Keduanya**.
- **🖨️ Cetak** membuka dialog cetak browser. Atur ukuran kertas **A5** (atau A4 dengan skala 100% / Actual size,
  lalu potong setengah) dan margin None. Jika dialog tidak muncul (mis. Safari/HP), pakai **Buka di tab baru** lalu Ctrl+P.

## Batasan
- Kuota Google API ±60 pembacaan/menit per pengguna; aplikasi men-cache data 20 detik, cukup untuk beberapa pengguna.
- Bila banyak orang menyimpan bersamaan, nomor tetap dicek ulang agar tidak ganda, tapi untuk beban tinggi
  gunakan database sungguhan (Postgres).

## Jalankan lokal
`pip install -r requirements.txt && streamlit run app.py` (tanpa Secrets → SQLite lokal).
