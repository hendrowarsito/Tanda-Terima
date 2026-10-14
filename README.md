# Tanda Terima (prototipe)

Pengganti OmniForm: penomoran otomatis, master klien/pengantar, item dinamis,
PDF tanda terima + amplop, riwayat, status pengiriman, dasbor "belum kembali", cadangan JSON.

## Deploy ke Streamlit Community Cloud
1. Unggah `app.py` dan `requirements.txt` ke repo GitHub.
2. share.streamlit.io → New app → pilih repo → main file `app.py`.
3. (Disarankan) *Settings → Secrets*:
```toml
app_password = "ganti-ini"
database_url = "postgresql://USER:PASS@HOST:5432/DB"   # Supabase/Neon, agar data tidak hilang
```
Tanpa `database_url` aplikasi memakai SQLite lokal yang **terhapus saat restart** — gunakan menu
Pengaturan → Unduh cadangan.

## Jalankan lokal
`pip install -r requirements.txt && streamlit run app.py`
