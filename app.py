"""Tanda Terima — prototipe pengganti OmniForm (Streamlit + Google Sheets / SQLite + ReportLab)."""
import base64, io, json, sqlite3
from datetime import datetime, timedelta
from pathlib import Path
from xml.sax.saxutils import escape
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
try:
    import pymupdf  # pratinjau PDF -> gambar
except ImportError:
    pymupdf = None
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Paragraph

st.set_page_config(page_title="Tanda Terima", page_icon="🧾", layout="wide")
WIB = ZoneInfo("Asia/Jakarta")
JENIS = {"PR": "Proposal", "INV": "Invoice", "KW": "Kwitansi", "LP": "Laporan", "UM": "Umum"}
STATUS = ["Draft", "Diserahkan", "Diterima", "Arsip"]
BLN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
       "September", "Oktober", "November", "Desember"]
KOP = {"nama": "SUWENDHO RINALDY DAN REKAN",
       "kiri": "KANTOR JASA PENILAI PUBLIK\nNomor Izin Usaha KJPP: 2.09.0059\nNomor Izin Cabang KJPP: 1138/KM.1/2017\nPenilai Properti dan Bisnis",
       "kanan": "Kantor Cabang Jakarta\nKomplek Kalibata Indah Blok K16-17\nJl. Rawajati Timur, Pancoran\nJakarta Selatan 12750\nT (021) 7970913 / 799-4521\nE ocky@srrkjpp.com\nWilayah Kerja: Seluruh Indonesia\nKantor Cabang: Jakarta (P/B), Bandung (P)"}
COLS = {"clients": ["id", "nama", "alamat1", "alamat2", "alamat3", "alamat4", "kode_pos", "kontak"],
        "couriers": ["id", "nama", "aktif"],
        "receipts": ["id", "jenis", "tahun", "seq", "nomor", "tanggal", "jam", "client_id", "up",
                     "courier_id", "status", "catatan", "items", "created_at", "penerima", "tgl_terima", "jam_terima"],
        "settings": ["k", "v"]}
INTS = {"id", "tahun", "seq", "client_id", "courier_id", "aktif"}


def secret(k, d=None):
    try:
        return st.secrets.get(k, d)
    except Exception:
        return d


# ---------- Penyimpanan: Google Sheets (utama) atau SQLite (cadangan) ----------
def clean(v):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if hasattr(v, "item"):
        v = v.item()
    return int(v) if isinstance(v, bool) else v


def cast(d, t):
    if not isinstance(d, pd.DataFrame):
        d = pd.DataFrame(columns=COLS[t])
    
    # Pastikan semua kolom mengikut skema COLS[t] wujud dalam DataFrame
    for c in COLS[t]:
        if c not in d.columns:
            d[c] = ""
            
    d = d.reindex(columns=COLS[t]).fillna("")
    for c in COLS[t]:
        if c in INTS:
            d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0).astype(int)
        else:
            d[c] = d[c].astype(str)
    return d.reset_index(drop=True)
    

class SqliteStore:
    def __init__(self, path):
        self.c = sqlite3.connect(path, check_same_thread=False)
        for t, cols in COLS.items():
            defs = ", ".join('"id" INTEGER PRIMARY KEY AUTOINCREMENT' if c == "id" else
                             '"k" TEXT PRIMARY KEY' if t == "settings" and c == "k" else f'"{c}"' for c in cols)
            self.c.execute(f"CREATE TABLE IF NOT EXISTS {t} ({defs})")
            have = {r[1] for r in self.c.execute(f"PRAGMA table_info({t})")}
            for col in cols:  # migrasi: tambahkan kolom baru pada database lama
                if col not in have:
                    self.c.execute(f'ALTER TABLE {t} ADD COLUMN "{col}"')
        self.c.commit()

    def read(self, t):
        return pd.read_sql(f"SELECT * FROM {t}", self.c)

    def insert(self, t, vals):
        ks = [k for k in COLS[t] if k in vals]
        cur = self.c.execute(f"INSERT INTO {t} ({','.join(chr(34) + k + chr(34) for k in ks)}) "
                             f"VALUES ({','.join('?' * len(ks))})", [clean(vals[k]) for k in ks])
        self.c.commit()
        return cur.lastrowid

    def update(self, t, key, vals):
        self.c.execute(f"UPDATE {t} SET {','.join(chr(34) + k + chr(34) + '=?' for k in vals)} "
                       f'WHERE "{COLS[t][0]}"=?', [clean(v) for v in vals.values()] + [key])
        self.c.commit()

    def delete(self, t, key):
        self.c.execute(f'DELETE FROM {t} WHERE "{COLS[t][0]}"=?', [key])
        self.c.commit()

    def replace(self, t, recs):
        self.c.execute(f"DELETE FROM {t}")
        for r in recs:
            self.insert(t, r)


class SheetsStore:
    """Tiap tabel = satu worksheet; baris 1 = judul kolom. Kolom pertama = kunci (id / k)."""

    def __init__(self, sid, info):
        import gspread
        self.gs = gspread
        self.sh = gspread.service_account_from_dict(dict(info)).open_by_key(sid)
        have = {w.title: w for w in self.sh.worksheets()}
        self.ws = {}
        for t, cols in COLS.items():
            w = have.get(t) or self.sh.add_worksheet(title=t, rows=1000, cols=len(cols))
            if w.row_values(1) != cols:  # sheet lama: tambah kolom baru di ujung kanan
                if w.col_count < len(cols):
                    w.resize(cols=len(cols))
                w.update(values=[cols], range_name="A1")
            self.ws[t] = w

    def read(self, t):
        n = len(COLS[t])
        rows = [(r + [""] * n)[:n] for r in self.ws[t].get_all_values()[1:] if r and r[0] != ""]
        return pd.DataFrame(rows, columns=COLS[t])

    def _row(self, t, key):
        return self.ws[t].col_values(1).index(str(key)) + 1

    def insert(self, t, vals):
        cols = COLS[t]
        if cols[0] == "id":
            nums = [int(x) for x in self.ws[t].col_values(1)[1:] if str(x).isdigit()]
            vals = {**vals, "id": max(nums, default=0) + 1}
        self.ws[t].append_row([clean(vals.get(c, "")) for c in cols], table_range="A1")
        return vals.get("id")

    def update(self, t, key, vals):
        r = self._row(t, key)
        self.ws[t].batch_update([{"range": self.gs.utils.rowcol_to_a1(r, COLS[t].index(k) + 1),
                                  "values": [[clean(v)]]} for k, v in vals.items()])

    def delete(self, t, key):
        self.ws[t].delete_rows(self._row(t, key))

    def replace(self, t, recs):
        self.ws[t].clear()
        self.ws[t].update(values=[COLS[t]] + [[clean(r.get(c, "")) for c in COLS[t]] for r in recs],
                          range_name="A1")


@st.cache_resource
def get_store():
    sid, info = secret("gsheet_id"), secret("gcp_service_account")
    return SheetsStore(sid, info) if sid and info else SqliteStore("tandaterima.db")


try:
    store = get_store()
except Exception as _e:  # tampilkan diagnosis (pesan asli Streamlit disembunyikan)
    st.error("Tidak bisa terhubung ke Google Sheets.")
    st.code(f"gsheet_id = {secret('gsheet_id')}\n"
            f"service account = {(secret('gcp_service_account') or {}).get('client_email')}\n"
            f"{type(_e).__name__}: {_e}")
    st.markdown(
        "**Cek berikut:**\n"
        "1. Bagikan (*Share*) Google Sheet ke email *service account* di atas dengan peran **Editor**.\n"
        "2. Di Google Cloud, aktifkan **Google Sheets API** dan **Google Drive API** untuk project-nya.\n"
        "3. `gsheet_id` = bagian URL antara `/d/` dan `/edit`.\n"
        "4. `private_key` di Secrets harus utuh, diapit tanda kutip, dengan `\\n` tetap seperti di file JSON.\n"
        "5. Mau mencoba dulu tanpa Sheets? Hapus `gsheet_id` dan `[gcp_service_account]` di Secrets (SQLite sementara).")
    st.stop()


@st.cache_data(ttl=20, show_spinner=False)
def _cached(t):
    return cast(store.read(t), t)


def D(t, fresh=False):
    """Baca tabel (di-cache 20 detik agar tidak kena batas kuota Google API)."""
    return cast(store.read(t), t) if fresh else _cached(t)


def W(op, t, *a):
    r = getattr(store, op)(t, *a)
    st.cache_data.clear()
    return r


def idn(iso):
    d = datetime.fromisoformat(iso)
    return f"{d.day} {BLN[d.month - 1]} {d.year}"


def get_kop():
    s = D("settings")
    return {**KOP, **dict(zip(s.k, s.v))}


def save_receipt(data):
    for _ in range(4):  # cek ulang setelah menulis: jika nomor bentrok dengan pengguna lain, ulangi
        d = D("receipts", True)
        m = d[(d.jenis == data["jenis"]) & (d.tahun == data["tahun"])].seq.max()
        seq = int(m) + 1 if pd.notna(m) else 1
        rid = W("insert", "receipts", {**data, "seq": seq, "nomor": f"{seq:04d}-{data['jenis']}"})
        d = D("receipts", True)
        dup = d[(d.jenis == data["jenis"]) & (d.tahun == data["tahun"]) & (d.seq == seq)]
        if len(dup) == 1 or dup.id.min() == rid:
            return rid
        W("delete", "receipts", rid)
    raise RuntimeError("Gagal membuat nomor unik, coba lagi.")


def sync(t, df, cols, ref_col, label):
    """Simpan hasil st.data_editor: tambah / ubah (hanya yang berubah) / hapus (diblokir bila sudah dipakai)."""
    cur = D(t, True).set_index("id")
    used = set(D("receipts", True)[ref_col]) if ref_col in D("receipts", True).columns else set()
    keep = set()

    for _, r in df.iterrows():
        # Menggunakan r.get(k, "") agar tidak memicu KeyError jika kolom tidak ditemukan
        vals = {k: clean(r.get(k, "")) for k in cols}
        
        # Jika kolom utama (misal 'nama') kosong, abaikan baris ini
        if not str(vals.get(cols[0], "")).strip():
            continue

        # Cek apakah ini baris baru (id kosong/NaN)
        row_id = r.get("id", None)
        if pd.isna(row_id) or row_id == "" or str(row_id).strip() == "":
            W("insert", t, vals)
            continue

        i = int(row_id)
        keep.add(i)
        
        # Perbarui jika ada perubahan data
        if i in cur.index:
            if any(str(cur.at[i, k]) != str(vals[k]) for k in cols if k in cur.columns):
                W("update", t, i, vals)

    # Hapus data yang dihapus dari editor (kecuali jika sudah pernah dipakai di tanda terima)
    for i in cur.index:
        if i in keep:
            continue
        if i in used:
            st.warning(f"{label} id {i} sudah dipakai di tanda terima, tidak dihapus (nonaktifkan saja).")
        else:
            W("delete", t, int(i))


def load_receipts():
    clients_df = D("clients")
    if "id" in clients_df.columns and "nama" in clients_df.columns:
        c = clients_df[["id", "nama"]].rename(columns={"id": "client_id", "nama": "klien"})
    else:
        c = pd.DataFrame(columns=["client_id", "klien"])

    couriers_df = D("couriers")
    if "id" in couriers_df.columns and "nama" in couriers_df.columns:
        k = couriers_df[["id", "nama"]].rename(columns={"id": "courier_id", "nama": "pengantar"})
    else:
        k = pd.DataFrame(columns=["courier_id", "pengantar"])

    receipts_df = D("receipts")
    if receipts_df.empty or "client_id" not in receipts_df.columns:
        return pd.DataFrame(columns=["id", "jenis", "tahun", "seq", "nomor", "tanggal", 
                                     "jam", "client_id", "up", "courier_id", "status", 
                                     "catatan", "items", "created_at", "klien", "pengantar"])

    return receipts_df.merge(c, on="client_id", how="left").merge(k, on="courier_id", how="left")



# ---------- PDF: setengah A4 (A5 landscape), meniru formulir OmniForm ----------
PAGE = (210 * mm, 297 * mm)  # A4: setengah atas = tanda terima, setengah bawah = amplop
HALF = 148.5  # mm dari tepi atas
BLUE = colors.HexColor("#1b1b8a")
COLW = [9.9, 90.2, 20.1, 49.8]  # mm (total 170)
LEFT, RIGHT = 20.2, 190.2
TOP_T, HEAD_H, ROW_H, MAX_BOTTOM = 68.6, 5.3, 5.5, 119.0  # mm dari tepi atas
LOGO_FILE = Path(globals().get("__file__", "app.py")).with_name("logo.png")  # opsional: logo asli
LOGO_B64 = """
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
"""  # logo KJPP (PNG); ganti dengan file logo.png bila perlu


def _logo():
    src = str(LOGO_FILE) if LOGO_FILE.exists() else io.BytesIO(base64.b64decode("".join(LOGO_B64.split())))
    return ImageReader(src)
_S_NAMA = 46.7 * mm / stringWidth(KOP["nama"], "Helvetica", 12.5) * 100
_S_BODY = 50.5 * mm / stringWidth("Nomor Izin Cabang KJPP: 1138/KM.1/2017", "Helvetica", 9) * 100


def _cond(c, txt, x, y, size, scale, maxw):
    """Teks 'condensed' (diperlebar-sempit horizontal) seperti kop aslinya."""
    sc = min(scale, maxw / max(stringWidth(txt, "Helvetica", size), 1) * 100)
    t = c.beginText(x, y)
    t.setFont("Helvetica", size)
    t.setHorizScale(sc)
    t.textOut(txt)
    t.setHorizScale(100)  # kembalikan agar teks berikutnya normal
    c.drawText(t)


def _fit(c, txt, x, y, font, size, maxw, align="l"):
    while size > 6 and stringWidth(txt, font, size) > maxw:
        size -= 0.25
    c.setFont(font, size)
    {"l": c.drawString, "c": c.drawCentredString, "r": c.drawRightString}[align](x, y, txt)
    return stringWidth(txt, font, size)


def _mix(c, parts, cx, y, size):
    x = cx - sum(stringWidth(t, f, size) for t, f in parts) / 2
    for t, f in parts:
        c.setFont(f, size)
        c.drawString(x, y, t)
        x += stringWidth(t, f, size)


def _label(c, x, y, a, b, size=10):
    c.setFont("Helvetica", size)
    c.drawString(x, y, a)
    c.setFont("Helvetica-Oblique", size)
    c.drawString(x, y - 4 * mm, b)


def _addr_lines(cl, up):
    """Maksimal 4 baris: nama + alamat (digabung bila kebanyakan) + U.p."""
    addr = [cl[k] for k in ("alamat1", "alamat2", "alamat3", "alamat4") if cl[k]]
    free = 3 - (1 if up else 0)
    if len(addr) > free:
        addr = addr[:free - 1] + [", ".join(addr[free - 1:])]
    return [cl["nama"]] + addr + ([f"U.p. {up}"] if up else [])


def _rows(items):
    base = ParagraphStyle("p", fontName="Helvetica", fontSize=8, leading=9.4)
    ctr = ParagraphStyle("pc", parent=base, alignment=1)
    out = []
    for n, it in enumerate(items, 1):
        txt = [escape(str(it.get(k) or "")).replace("\n", "<br/>") for k in ("Uraian", "Jumlah", "Keterangan")]
        cells = [Paragraph(f"{n}.", ctr), Paragraph(txt[0], base), Paragraph(txt[1], ctr), Paragraph(txt[2], base)]
        for p, w in zip(cells, COLW):
            p.wrap((w - 3) * mm, 1000)
        out.append((cells, max(ROW_H * mm, max(p.height for p in cells) + 1.8 * mm)))
    return out


def _paginate(rows, limit):
    pages, cur, used = [], [], 0
    for rw in rows:
        if cur and used + rw[1] > limit:
            pages.append(cur)
            cur, used = [], 0
        cur.append(rw)
        used += rw[1]
    return pages + [cur]


def _form(c, r, cl, kurir, kop, rows, pad, pi, np_):
    H = PAGE[1]
    T = lambda m: H - m * mm
    X = lambda m: m * mm
    # logo
    lx, ly, lh = 16.2, 6.5, 22
    img = _logo()
    iw, ih = img.getSize()
    c.drawImage(img, X(lx), T(ly + lh), lh * iw / ih * mm, lh * mm, mask="auto")
    # kop
    c.setFillColor(BLUE)
    _cond(c, kop["nama"], X(39), T(12.8), 12.5, _S_NAMA, 62 * mm)
    y = 16.5
    for ln in kop["kiri"].split("\n"):
        _cond(c, ln, X(39), T(y), 9, _S_BODY, 52 * mm)
        y += 3.4
    y = 6.5
    for n, ln in enumerate(kop["kanan"].split("\n")):
        _fit(c, ln, X(91.6), T(y), "Helvetica-Bold" if n == 0 else "Helvetica", 7, 46.5 * mm)
        y += 2.85
    # nomor
    c.setFillColor(colors.black)
    c.setStrokeColor(colors.black)
    c.setFont("Helvetica", 10)
    c.drawString(X(160), T(14.2), "No.")
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(X(160), T(18.4), "Number")
    c.setLineWidth(.5)
    c.line(X(166), T(15.2), X(185), T(15.2))
    _fit(c, r["nomor"], X(175.5), T(14.4), "Helvetica", 11, 18.5 * mm, "c")
    if np_ > 1:
        c.setFont("Helvetica", 7.5)
        c.drawString(X(160), T(22.6), f"Hal. {pi}/{np_}")
    # garis tebal + kotak judul
    c.setLineWidth(2)
    c.line(X(19.5), T(31.9), X(190.3), T(31.9))
    c.setFillColor(colors.white)
    c.setLineWidth(1.3)
    c.roundRect(X(139.2), T(37.8), 36.5 * mm, 11.5 * mm, 2 * mm, stroke=1, fill=1)
    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 11.5)
    c.drawCentredString(X(157.45), T(31.1), "TANDA TERIMA")
    cs = 2.6
    fw = stringWidth("RECEIPT", "Helvetica-Oblique", 9) + cs * 6
    t = c.beginText(X(157.45) - fw / 2, T(35.3))
    t.setFont("Helvetica-Oblique", 9)
    t.setCharSpace(cs)
    t.textOut("RECEIPT")
    t.setCharSpace(0)  # kembalikan jarak huruf
    c.drawText(t)
    # Kepada
    c.setFont("Helvetica", 10)
    c.drawString(X(LEFT), T(38.8), "Kepada:")
    c.setLineWidth(.5)
    c.line(X(LEFT), T(39.4), X(LEFT) + stringWidth("Kepada:", "Helvetica", 10), T(39.4))
    c.setFont("Helvetica-Oblique", 10)
    c.drawString(X(LEFT), T(42.9), "To")
    lines = _addr_lines(cl, r["up"])
    for i, ys in enumerate((39.9, 44.4, 49.2, 53.9)):
        c.setStrokeColor(colors.HexColor("#555555"))
        c.setLineWidth(.4)
        c.line(X(34.2), T(ys), X(117.5), T(ys))
        if i < len(lines):
            _fit(c, lines[i], X(35), T(ys - .8), "Helvetica-Bold" if i == 0 else "Helvetica", 9.5, 81 * mm)
    c.setStrokeColor(colors.black)
    # Tgl & Jam
    _label(c, X(131), T(45.4), "Tgl.", "Date")
    c.setLineWidth(.4)
    c.line(X(139.8), T(46.2), X(180.2), T(46.2))
    _fit(c, idn(r["tanggal"]), X(141), T(45.5), "Helvetica", 9.5, 38 * mm)
    _label(c, X(131.3), T(56.1), "Jam", "Hour")
    c.line(X(140), T(56.8), X(160), T(56.8))
    c.setFont("Helvetica", 9.5)
    c.drawString(X(141), T(56.0), r["jam"])
    _label(c, X(LEFT), T(61.9), "Telah diterima/diserahkan:", "Received/delivered:")
    # tabel
    x0 = [LEFT]
    for w in COLW:
        x0.append(x0[-1] + w)
    bottom = TOP_T + HEAD_H + sum(h for _, h in rows) / mm + pad * ROW_H
    c.setLineWidth(1.5)
    c.rect(X(LEFT), T(bottom), (RIGHT - LEFT) * mm, (bottom - TOP_T) * mm)
    for xv in x0[1:-1]:
        c.line(X(xv), T(TOP_T), X(xv), T(bottom))
    y = TOP_T + HEAD_H
    c.line(X(LEFT), T(y), X(RIGHT), T(y))
    O, I = "Helvetica", "Helvetica-Oblique"
    for i, parts in enumerate([[("No.", O)], [("Uraian/", O), ("Description", I)],
                               [("Jumlah/", O), ("Qty", I)], [("Keterangan/", O), ("Remarks", I)]]):
        _mix(c, parts, X((x0[i] + x0[i + 1]) / 2), T(TOP_T + 3.9), 9)
    for cells, h in rows:
        for i, p in enumerate(cells):
            p.drawOn(c, X(x0[i]) + 1.5 * mm, T(y) - (h - p.height) / 2 - p.height)
        y += h / mm
        c.line(X(LEFT), T(y), X(RIGHT), T(y))
    for _ in range(pad):
        y += ROW_H
        c.line(X(LEFT), T(y), X(RIGHT), T(y))
    # tanda tangan
    sh = max(0, bottom - 112.5)
    _label(c, X(LEFT), T(117.8 + sh), "Diterima oleh:", "Received by")
    _label(c, X(138.6), T(117.8 + sh), "Diserahkan oleh:", "Delivered by")
    c.setLineWidth(.6)
    c.line(X(20), T(142.3), X(59), T(142.3))
    c.line(X(138), T(142.3), X(167.7), T(142.3))
    if kurir:
        c.setFont("Helvetica", 8)
        c.drawCentredString(X(152.85), T(145.8), kurir)
    pen = str(r.get("penerima", "") or "")
    if pen:
        tg = str(r.get("tgl_terima", "") or "")
        _fit(c, f"{pen} — {idn(tg)}, {r.get('jam_terima', '')}" if tg else pen, X(20), T(145.8), "Helvetica", 8, 100 * mm)
    if r["catatan"]:
        _fit(c, f"Catatan: {r['catatan']}", X(63), T(141.3), "Helvetica-Oblique", 7.5, 72 * mm)


def _envelope(c, cl, up, cut=False):
    """Amplop di setengah bawah A4: bingkai, dan garis bawah hanya pada baris alamat terakhir."""
    T = lambda m: PAGE[1] - m * mm
    X = lambda m: m * mm
    c.setFillColor(colors.black)
    if cut:  # garis potong antara tanda terima dan amplop
        c.setStrokeColor(colors.HexColor("#999999"))
        c.setLineWidth(.4)
        c.setDash(2, 2)
        c.line(X(8), T(HALF), X(202), T(HALF))
        c.setDash()
    c.setStrokeColor(colors.black)
    c.setLineWidth(1.2)
    c.rect(X(16), T(245), 100 * mm, 70 * mm)  # bingkai amplop
    x, maxw = 30, 150 * mm
    c.setFont("Helvetica", 13)
    c.drawString(X(x), T(186), "Kepada Yth.")
    _fit(c, cl["nama"], X(x), T(197), "Helvetica-Bold", 12, maxw)
    y = 197
    addr = [cl[k] for k in ("alamat1", "alamat2", "alamat3", "alamat4") if cl[k]]
    for i, ln in enumerate(addr):
        y += 8
        w = _fit(c, ln, X(x), T(y), "Helvetica", 10, maxw)
        if i == len(addr) - 1:  # underline hanya baris alamat terakhir
            c.setLineWidth(.8)
            c.line(X(x), T(y + 1.6), X(x) + w, T(y + 1.6))
    if up:
        _fit(c, f"U.p. : {up}", X(x), T(y + 14), "Helvetica-Bold", 10, maxw)


def make_pdf(r, cl, kurir, kop, items, mode="semua"):
    """Selalu kertas A4. mode: 'semua' = tanda terima (atas) + amplop (bawah), 'tt' = tanda terima saja,
    'amplop' = amplop saja (posisi tetap, sehingga separuh kertas yang kosong bisa dipakai lagi)."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=PAGE)
    pages = [[]] if mode == "amplop" else _paginate(_rows(items), (MAX_BOTTOM - TOP_T - HEAD_H) * mm)
    for i, rows in enumerate(pages, 1):
        if mode != "amplop":
            used, pad = sum(h for _, h in rows) / mm, 0
            while len(rows) + pad < 7 and used + (pad + 1) * ROW_H <= 7 * ROW_H + .2:
                pad += 1
            _form(c, r, cl, kurir, kop, rows, pad, i, len(pages))
        if i == 1 and mode != "tt":
            _envelope(c, cl, r["up"], cut=(mode == "semua"))
        c.showPage()
    c.save()
    return buf.getvalue()


def pdf_for(rid, mode="semua"):
    d = D("receipts")
    r = d[d.id == rid].iloc[0]
    cl = D("clients")
    cl = cl[cl.id == r.client_id].iloc[0]
    k = D("couriers")
    k = k[k.id == r.courier_id]
    return make_pdf(r, cl, k.nama.iloc[0] if len(k) else "", get_kop(), json.loads(r["items"]), mode)


# ---------- Pratinjau & cetak ----------
PRINT_HTML = r"""
<style>
 body{margin:0;font-family:"Source Sans Pro",system-ui,sans-serif}
 button{padding:.45rem .8rem;margin:0 .35rem .35rem 0;border:1px solid #bbb;border-radius:.5rem;background:#fff;cursor:pointer;font-size:14px}
 button.p{background:#ff4b4b;border-color:#ff4b4b;color:#fff}
 #m{font-size:13px;color:#b00020}
</style>
<button class="p" id="p">🖨️ Cetak</button><button id="o">📄 Buka PDF</button><span id="m"></span>
<script>
const PAGES=__PAGES__, PDF="__PDF__";
const $=id=>document.getElementById(id), msg=t=>{$("m").textContent=t;};
const css="@page{size:A4 portrait;margin:0}html,body{margin:0;padding:0;background:#fff}"+
 "img{display:block;width:99.6%;height:auto;page-break-after:always;break-after:page}"+
 "img:last-of-type{page-break-after:auto;break-after:auto}.bar{display:none}"+
 "@media screen{html,body{background:#777}.bar{display:block;position:sticky;top:0;padding:8px;background:#333;text-align:center}"+
 ".bar button{padding:8px 18px;font-size:15px;cursor:pointer}img{width:min(210mm,96%);margin:12px auto;box-shadow:0 0 8px #000}}";
const bar='<div class="bar"><button onclick="window.print()">🖨️ Cetak</button></div>';
const AUTO='<script>window.onload=function(){setTimeout(function(){window.focus();window.print()},400)};'+
 'window.onafterprint=function(){window.close()}<\/script>';
const html='<!doctype html><html><head><meta charset="utf-8"><title>Tanda Terima</title><style>'+css+'</style></head><body>'+
 bar+PAGES.map(p=>'<img src="data:image/png;base64,'+p+'">').join("")+AUTO+'</body></html>';
const bin=atob(PDF),arr=new Uint8Array(bin.length);
for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);
const url=URL.createObjectURL(new Blob([arr],{type:"application/pdf"}));
if(!PAGES.length)$("p").style.display="none";
$("o").onclick=()=>{if(!window.open(url,"_blank"))msg("Pop-up diblokir browser.");};
$("p").onclick=()=>{
  const w=window.open("","_blank");
  if(!w){msg("Pop-up diblokir: izinkan pop-up untuk situs ini, atau pakai Unduh PDF.");return;}
  msg("");w.document.open();w.document.write(html);w.document.close();
};
</script>"""


def print_bar(pdf):
    """Cetak lewat jendela baru berisi gambar halaman A4 (bukan pemutar PDF / iframe), agar andal di Safari/Chrome/Firefox."""
    pages = [base64.b64encode(p).decode() for p in render_pages(pdf, 200)]
    html = PRINT_HTML.replace("__PAGES__", json.dumps(pages)).replace("__PDF__", base64.b64encode(pdf).decode())
    if hasattr(st, "iframe"):
        st.iframe(html, height=64)
    else:
        import streamlit.components.v1 as comp
        comp.html(html, height=64)


@st.cache_data(show_spinner=False, max_entries=30)
def render_pages(pdf, dpi=120):
    if pymupdf is None:
        return []
    with pymupdf.open(stream=pdf, filetype="pdf") as doc:
        return [p.get_pixmap(dpi=dpi).tobytes("png") for p in doc]


def preview(pdf):
    imgs = render_pages(pdf)
    if not imgs:
        st.info("Pratinjau butuh paket `pymupdf` (lihat requirements.txt).")
        return
    if len(imgs) == 1:
        st.image(imgs[0], width=640)
    else:
        for tab, img in zip(st.tabs([f"Lembar {i}/{len(imgs)}" for i in range(1, len(imgs) + 1)]), imgs):
            tab.image(img, width=640)


def show_receipt(rid, key):
    opts = {"Tanda terima + amplop": "semua", "Tanda terima saja": "tt", "Amplop saja": "amplop"}
    pick = st.segmented_control("Yang dicetak / diunduh", list(opts), default="Tanda terima + amplop",
                                key=f"mode2_{key}") or "Tanda terima + amplop"
    pdf = pdf_for(rid, opts[pick])
    d = D("receipts")
    no = d.loc[d.id == rid, "nomor"].iloc[0]
    a, b = st.columns([4, 1.3])
    with a:
        print_bar(pdf)
    b.download_button("⬇️ Unduh PDF", pdf, f"{no}.pdf", "application/pdf", key=f"dl_{key}")
    st.caption("Kertas **A4**: setengah atas tanda terima, setengah bawah amplop (garis putus-putus = tempat memotong). "
               "**Cetak** membuka jendela baru berisi halaman siap cetak (izinkan pop-up bila diminta); jika dialog tidak "
               "muncul sendiri, tekan tombol Cetak di jendela itu. Di dialog: ukuran **A4**, skala **100%**, matikan "
               "*header & footer*. Cadangan: **Buka PDF** lalu Cmd/Ctrl+P.")
    preview(pdf)


def next_nomor(jenis, tahun):
    d = D("receipts")
    m = d[(d.jenis == jenis) & (d.tahun == tahun)].seq.max()
    return f"{(int(m) + 1 if pd.notna(m) else 1):04d}-{jenis}"


# ---------- Halaman ----------
def page_buat():
    # Pastikan DataFrame clients tidak kosong dan mempunyai kolom 'nama' sebelum sort
    cl = D("clients")
    if "nama" in cl.columns and not cl.empty:
        cl = cl.sort_values("nama")
    else:
        cl = pd.DataFrame(columns=COLS["clients"])

    cr = D("couriers")
    if "aktif" in cr.columns:
        cr = cr[cr.aktif == 1]
    else:
        cr = pd.DataFrame(columns=COLS["couriers"])

    if cl.empty or cr.empty:
        st.info("Isi dulu menu **Klien & Pengantar** (atau pakai data contoh di **Pengaturan**).")
        return
        
    now = datetime.now(WIB)
    n = st.session_state.get("n", 0)
    st.session_state.setdefault("b_tgl", now.date())
    st.session_state.setdefault("b_jam", now.time().replace(second=0, microsecond=0))
    a, b, c3, d = st.columns(4)
    jenis = a.selectbox("Jenis", list(JENIS), format_func=lambda k: f"{k} – {JENIS[k]}")
    tgl = b.date_input("Tanggal", key="b_tgl")
    jam = c3.time_input("Jam", key="b_jam")
    kur = d.selectbox("Pengantar", cr.id.tolist(), format_func=lambda i: cr.set_index("id").nama[i])
    nm = cl.set_index("id").nama
    cid = st.selectbox("Kepada", cl.id.tolist(), format_func=lambda i: nm[i])
    row = cl[cl.id == cid].iloc[0]
    st.caption(" · ".join(x for x in [row.alamat1, row.alamat2, row.alamat3, row.alamat4] if x))
    kontak = [x.strip() for x in row.kontak.split("\n") if x.strip()]
    up = st.selectbox("U.p. (pilih atau ketik baru)", kontak, index=0 if kontak else None,
                      accept_new_options=True)
    st.write("**Daftar dokumen**")
    items = st.data_editor(pd.DataFrame([{"Uraian": "", "Jumlah": "1 asli", "Keterangan": ""}]),
                           num_rows="dynamic", width="stretch", key=f"items{n}")
    catatan = st.text_input("Catatan (opsional)")
    its = [r for r in items.fillna("").to_dict("records") if str(r["Uraian"]).strip()]

    st.subheader("Pratinjau")
    rprev = dict(nomor=next_nomor(jenis, tgl.year), tanggal=tgl.isoformat(), jam=jam.strftime("%H:%M"),
                 up=up or "", catatan=catatan)
    preview(make_pdf(rprev, row, cr.set_index("id").nama[kur], get_kop(), its, "semua"))
    st.caption("Nomor final ditetapkan saat disimpan. Tombol cetak muncul setelah disimpan.")
    if st.button("💾 Simpan", type="primary"):
        if not its:
            st.error("Isi minimal satu baris dokumen.")
        else:
            rid = save_receipt(dict(jenis=jenis, tahun=tgl.year, tanggal=tgl.isoformat(), jam=jam.strftime("%H:%M"),
                                    client_id=int(cid), up=up or "", courier_id=int(kur), status="Draft",
                                    catatan=catatan, items=json.dumps(its, ensure_ascii=False),
                                    created_at=now.isoformat(timespec="seconds")))
            st.session_state["last"] = rid
            st.session_state["n"] = n + 1
            st.rerun()
    
    st.header("Buat tanda terima")
    if "last" in st.session_state:
        d = D("receipts")
        m = d[d.id == st.session_state["last"]]
        if m.empty:
            st.session_state.pop("last")
        else:
            with st.container(border=True):
                st.success(f"Tanda terima **{m.nomor.iloc[0]}** tersimpan. Siap dicetak:")
                show_receipt(int(m.id.iloc[0]), "last")
                if st.button("✖ Tutup & buat baru"):
                    st.session_state.pop("last")
                    st.rerun()


def _parse_d(x, default):
    try:
        return datetime.fromisoformat(x).date()
    except Exception:
        return default


def _parse_t(x, default):
    try:
        return datetime.strptime(x, "%H:%M").time()
    except Exception:
        return default


def page_riwayat():
    st.header("Riwayat")
    df = load_receipts()
    if df.empty:
        st.info("Belum ada tanda terima.")
        return
    a, b, c3 = st.columns([2, 1, 1])
    s = a.text_input("Cari (klien / nomor / isi dokumen)")
    jf = b.multiselect("Jenis", list(JENIS))
    sf = c3.multiselect("Status", STATUS)
    v = df.sort_values(["tahun", "seq"], ascending=False)
    if s:
        hay = (v.klien.fillna("") + v.nomor + v["items"]).str.lower()
        v = v[hay.str.contains(s.lower(), regex=False)]
    if jf:
        v = v[v.jenis.isin(jf)]
    if sf:
        v = v[v.status.isin(sf)]
    st.dataframe(v[["nomor", "jenis", "tanggal", "jam", "klien", "pengantar", "status", "penerima", "tgl_terima",
                    "jam_terima"]], hide_index=True, width="stretch")
    if v.empty:
        return
    st.divider()
    lab = {r.id: f"{r.nomor} · {r.tanggal} · {r.klien}" for r in v.itertuples()}
    rid = st.selectbox("Pilih tanda terima", list(lab), format_func=lab.get)
    r = v[v.id == rid].iloc[0]
    a, b = st.columns(2)
    ns = a.selectbox("Status", STATUS, index=STATUS.index(r.status) if r.status in STATUS else 0, key=f"st_{rid}")
    now = datetime.now(WIB)
    if ns == "Diterima":  # wajib isi penerima, tanggal, dan jam diterima
        pen = a.text_input("Penerima *", value=r.penerima, key=f"pen_{rid}")
        c1, c2 = a.columns(2)
        tg = c1.date_input("Tanggal diterima", value=_parse_d(r.tgl_terima, now.date()), key=f"tgt_{rid}")
        jm = c2.time_input("Jam diterima", value=_parse_t(r.jam_terima, now.time().replace(second=0, microsecond=0)),
                           key=f"jmt_{rid}")
    if a.button("Perbarui status"):
        if ns == "Diterima" and not pen.strip():
            a.error("Isi nama penerima.")
        else:
            vals = {"status": ns}
            if ns == "Diterima":
                vals.update(penerima=pen.strip(), tgl_terima=tg.isoformat(), jam_terima=jm.strftime("%H:%M"))
            W("update", "receipts", int(rid), vals)
            st.rerun()
    if b.checkbox("Izinkan hapus") and b.button("🗑️ Hapus"):
        W("delete", "receipts", int(rid))
        st.rerun()
    show_receipt(int(rid), f"h{rid}")

def page_master():
    st.header("Klien & Pengantar")
    st.subheader("Klien")
    st.caption("Kolom *kontak*: satu nama (U.p.) per baris. Hapus baris dengan memilihnya lalu tekan Delete.")
    
    clients_df = D("clients")
    # Pastikan semua kolom terdefinisi
    for c in COLS["clients"]:
        if c not in clients_df.columns:
            clients_df[c] = ""

    if "nama" in clients_df.columns and not clients_df.empty:
        clients_df = clients_df.sort_values("nama")

    ed = st.data_editor(
        clients_df,
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key="ce",
        disabled=["id"],
        column_config={"kontak": st.column_config.TextColumn(width="large")}
    )
    
    if st.button("Simpan klien"):
        sync("clients", ed, ["nama", "kode", "alamat1", "alamat2", "alamat3", "alamat4", "kontak"],
             "client_id", "Klien")
        st.success("Data klien berhasil disimpan!")
        st.rerun()

    st.subheader("Pengantar")
    kdf = D("couriers")
    for c in COLS["couriers"]:
        if c not in kdf.columns:
            kdf[c] = ""

    if "nama" in kdf.columns and not kdf.empty:
        kdf = kdf.sort_values("nama")
    
    if "aktif" in kdf.columns:
        kdf["aktif"] = kdf["aktif"].astype(bool)

    ek = st.data_editor(
        kdf,
        num_rows="dynamic",
        width="stretch",
        hide_index=True,
        key="ke",
        disabled=["id"],
        column_config={"aktif": st.column_config.CheckboxColumn()}
    )
    
    if st.button("Simpan pengantar"):
        sync("couriers", ek.fillna({"aktif": True}), ["nama", "aktif"], "courier_id", "Pengantar")
        st.success("Data pengantar berhasil disimpan!")
        st.rerun()


def page_setelan():
    st.header("Pengaturan")
    if isinstance(store, SheetsStore):
        st.success("Penyimpanan: **Google Sheets** (data tetap).")
        st.link_button("Buka Google Sheet", store.sh.url)
    else:
        st.warning("Penyimpanan: **SQLite sementara** — data hilang saat aplikasi restart di Streamlit Cloud. "
                   "Isi `gsheet_id` dan `[gcp_service_account]` di Secrets untuk memakai Google Sheets.")
    st.subheader("Kop surat")
    kop = get_kop()
    nama = st.text_input("Nama KJPP", kop["nama"])
    kiri = st.text_area("Blok kiri (di bawah nama)", kop["kiri"], height=110)
    kanan = st.text_area("Blok kanan (alamat cabang)", kop["kanan"], height=130)
    if st.button("Simpan kop"):
        have = set(D("settings", True).k)
        for k, v in dict(nama=nama, kiri=kiri, kanan=kanan).items():
            W("update", "settings", k, {"v": v}) if k in have else W("insert", "settings", {"k": k, "v": v})
        st.success("Tersimpan.")
    st.subheader("Cadangan data")
    data = {t: D(t, True).to_dict("records") for t in COLS}
    st.download_button("⬇️ Unduh cadangan (JSON)", json.dumps(data, ensure_ascii=False, default=str, indent=1),
                       f"cadangan-tanda-terima-{datetime.now(WIB):%Y%m%d}.json", "application/json")
    up = st.file_uploader("Pulihkan dari cadangan (menimpa semua data)", type="json")
    if up and st.button("Pulihkan sekarang"):
        d = json.load(up)
        for t in COLS:
            W("replace", t, d.get(t, []))
        st.success("Data dipulihkan.")
    st.subheader("Data contoh")
    if st.button("Isi data contoh (fiktif)"):
        W("insert", "clients", dict(nama="PT CONTOH ABADI Tbk", kode="CABD", alamat1="Menara Contoh Lt. 10",
                                    alamat2="Jl. Jenderal Sudirman Kav. 1", alamat3="Jakarta 10220", alamat4="",
                                    kontak="Bpk. Budi\nIbu Sari"))
        for n in ("Pengantar A", "Pengantar B"):
            W("insert", "couriers", {"nama": n, "aktif": 1})
        st.success("Data contoh ditambahkan.")


# ---------- Main ----------
pw = secret("app_password")
if pw and not st.session_state.get("ok"):
    st.title("🧾 Tanda Terima")
    if st.text_input("Kata sandi", type="password") == pw:
        st.session_state["ok"] = True
        st.rerun()
    st.stop()

PAGES = {"🧾 Buat": page_buat, "📋 Riwayat": page_riwayat,
         "👥 Klien & Pengantar": page_master, "⚙️ Pengaturan": page_setelan}
st.sidebar.title("Tanda Terima")
PAGES[st.sidebar.radio("Menu", list(PAGES), label_visibility="collapsed")]()
