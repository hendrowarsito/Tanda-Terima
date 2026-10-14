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
from reportlab.pdfgen import canvas
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.platypus import Paragraph

st.set_page_config(page_title="Tanda Terima", page_icon="🧾", layout="wide")
WIB = ZoneInfo("Asia/Jakarta")
JENIS = {"PR": "Proposal", "INV": "Invoice", "KW": "Kwitansi", "LP": "Laporan"}
STATUS = ["Draft", "Diserahkan", "Diterima", "Arsip"]
BLN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
       "September", "Oktober", "November", "Desember"]
KOP = {"nama": "SUWENDHO RINALDY & REKAN",
       "kiri": "KANTOR JASA PENILAI PUBLIK\nNomor Izin Usaha KJPP: 2.09.0059\nNomor Izin Cabang KJPP: 1138/KM.1/2017\nPenilai Properti dan Bisnis",
       "kanan": "Kantor Cabang Jakarta\nKomplek Kalibata Indah Blok K16-17\nJl. Rawajati Timur, Pancoran\nJakarta Selatan 12750\nT (021) 7970913 / 799-4521\nE ocky@srr.co.id\nWilayah Kerja: Seluruh Indonesia\nKantor Cabang: Jakarta (P/B), Bandung (P)"}
COLS = {"clients": ["id", "nama", "kode", "alamat1", "alamat2", "alamat3", "alamat4", "kontak"],
        "couriers": ["id", "nama", "aktif"],
        "receipts": ["id", "jenis", "tahun", "seq", "nomor", "tanggal", "jam", "client_id", "up",
                     "courier_id", "status", "catatan", "items", "created_at"],
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
            # Pastikan header kolom sesuai
            headers = w.row_values(1)
            if headers != cols:
                w.update(range_name="A1", values=[cols])
            self.ws[t] = w

    def read(self, t):
        n = len(COLS[t])
        rows = [(r + [""] * n)[:n] for r in self.ws[t].get_all_values()[1:] if r and r[0] != ""]
        return pd.DataFrame(rows, columns=COLS[t])

    def _row(self, t, key):
        col_vals = [str(x) for x in self.ws[t].col_values(1)]
        str_key = str(key)
        if str_key in col_vals:
            return col_vals.index(str_key) + 1
        raise KeyError(f"Key {key} tidak ditemukan di tabel {t}")

    def insert(self, t, vals):
        cols = COLS[t]
        if cols[0] == "id":
            nums = [int(x) for x in self.ws[t].col_values(1)[1:] if str(x).isdigit()]
            vals = {**vals, "id": max(nums, default=0) + 1}
        # Hapus parameter table_range="A1" agar append_row berjalan normal di semua versi gspread
        row_data = [clean(vals.get(c, "")) for c in cols]
        self.ws[t].append_row(row_data, value_input_option="USER_ENTERED")
        return vals.get("id")

    def update(self, t, key, vals):
        r = self._row(t, key)
        updates = []
        for k, v in vals.items():
            col_idx = COLS[t].index(k) + 1
            cell_range = self.gs.utils.rowcol_to_a1(r, col_idx)
            updates.append({"range": cell_range, "values": [[clean(v)]]})
        self.ws[t].batch_update(updates)

    def delete(self, t, key):
        r = self._row(t, key)
        self.ws[t].delete_rows(r)

    def replace(self, t, recs):
        self.ws[t].clear()
        data = [COLS[t]] + [[clean(r.get(c, "")) for c in COLS[t]] for r in recs]
        self.ws[t].update(range_name="A1", values=data)


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
    try:
        d = datetime.fromisoformat(str(iso)[:10])
    except ValueError:
        return str(iso)
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
    used = set(D("receipts", True)[ref_col])
    keep = set()
    for _, r in df.iterrows():
        vals = {k: clean(r[k]) for k in cols}
        if not str(vals[cols[0]]).strip():
            continue
        if pd.isna(r["id"]):
            W("insert", t, vals)
            continue
        i = int(r["id"])
        keep.add(i)
        if i in cur.index and any(str(cur.at[i, k]) != str(vals[k]) for k in cols):
            W("update", t, i, vals)
    for i in cur.index:
        if i in keep:
            continue
        if i in used:
            st.warning(f"{label} id {i} sudah dipakai di tanda terima, tidak dihapus (nonaktifkan saja).")
        else:
            W("delete", t, int(i))


def load_receipts():
    c = D("clients")[["id", "nama"]].rename(columns={"id": "client_id", "nama": "klien"})
    k = D("couriers")[["id", "nama"]].rename(columns={"id": "courier_id", "nama": "pengantar"})
    return D("receipts").merge(c, on="client_id", how="left").merge(k, on="courier_id", how="left")


# ---------- PDF: setengah A4 (A5 landscape), meniru formulir OmniForm ----------
PAGE = (210 * mm, 148.5 * mm)
BLUE = colors.HexColor("#1b1b8a")
NAVY = colors.HexColor("#301d6d")
COLW = [9.9, 90.2, 20.1, 49.8]  # mm (total 170)
LEFT, RIGHT = 20.2, 190.2
TOP_T, HEAD_H, ROW_H, MAX_BOTTOM = 68.6, 5.3, 5.5, 119.0  # mm dari tepi atas
LOGO_FILE = Path(globals().get("__file__", "app.py")).with_name("logo.png")  # opsional: logo asli
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
    lx, ly, ls = 16.2, 6.5, 22
    if LOGO_FILE.exists():
        c.drawImage(str(LOGO_FILE), X(lx), T(ly + ls), ls * mm, ls * mm, preserveAspectRatio=True, mask="auto")
    else:
        c.setFillColor(NAVY)
        c.rect(X(lx), T(ly + ls), ls * mm, ls * mm, stroke=0, fill=1)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 22)
        c.drawRightString(X(lx + .926 * ls), T(ly + .667 * ls), "SR")
        c.setFont("Helvetica-Bold", 19.5)
        c.drawRightString(X(lx + .937 * ls), T(ly + .958 * ls), "R")
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
    if r["catatan"]:
        _fit(c, f"Catatan: {r['catatan']}", X(63), T(141.3), "Helvetica-Oblique", 7.5, 72 * mm)


def _envelope(c, cl, up):
    T = lambda m: PAGE[1] - m * mm
    x = 35 * mm
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 12)
    c.drawString(x, T(50), "Kepada Yth.")
    c.setFont("Helvetica-Bold", 15)
    c.drawString(x, T(60), cl["nama"])
    y = 60
    c.setFont("Helvetica", 12)
    for k in ("alamat1", "alamat2", "alamat3", "alamat4"):
        if cl[k]:
            y += 6.5
            c.drawString(x, T(y), cl[k])
    if up:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(x, T(y + 10), f"U.p. : {up}")


def _items(raw):
    try:
        data = json.loads(raw) if isinstance(raw, str) else raw
    except (TypeError, json.JSONDecodeError):
        return []
    if not isinstance(data, list):
        return []
    out = []
    for it in data:
        if isinstance(it, dict):
            out.append({"Uraian": str(it.get("Uraian") or ""),
                        "Jumlah": str(it.get("Jumlah") or ""),
                        "Keterangan": str(it.get("Keterangan") or "")})
    return out


def _client_of(cid):
    cl = D("clients")
    hit = cl[cl.id == cid]
    if hit.empty:
        return pd.Series({k: "" for k in COLS["clients"]})
    return hit.iloc[0]


def make_pdf(r, cl, kurir, kop, items, mode="semua"):
    """mode: 'tt' = tanda terima saja, 'amplop' = amplop saja, 'semua' = keduanya."""
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=PAGE)
    if mode != "amplop":
        pages = _paginate(_rows(items), (MAX_BOTTOM - TOP_T - HEAD_H) * mm) or [[]]
        for i, rows in enumerate(pages, 1):
            used, pad = sum(h for _, h in rows) / mm, 0
            while len(rows) + pad < 7 and used + (pad + 1) * ROW_H <= 7 * ROW_H + .2:
                pad += 1
            _form(c, r, cl, kurir, kop, rows, pad, i, len(pages))
            # showPage setelah halaman terakhir membuat halaman kosong. Jangan panggil bila ini halaman akhir.
            if i < len(pages) or mode != "tt":
                c.showPage()
    if mode != "tt":
        _envelope(c, cl, r["up"])
    c.save()
    return buf.getvalue()


def pdf_for(rid, mode="semua"):
    d = D("receipts")
    hit = d[d.id == rid]
    if hit.empty:
        raise RuntimeError("Tanda terima tidak ditemukan.")
    r = hit.iloc[0]
    k = D("couriers")
    k = k[k.id == r.courier_id]
    return make_pdf(r, _client_of(r.client_id), k.nama.iloc[0] if len(k) else "",
                    get_kop(), _items(r["items"]), mode)


# ---------- Pratinjau & cetak ----------
PRINT_HTML = """
<style>
 body{margin:0;font-family:"Source Sans Pro",system-ui,sans-serif}
 button{padding:.45rem .9rem;margin-right:.4rem;border:1px solid #bbb;border-radius:.5rem;background:#fff;cursor:pointer;font-size:14px}
 button.p{background:#ff4b4b;border-color:#ff4b4b;color:#fff}
</style>
<button class="p" id="p">🖨️ Cetak</button><button id="o">↗️ Buka di tab baru</button>
<script>
const bin=atob("__B64__"),arr=new Uint8Array(bin.length);
for(let i=0;i<bin.length;i++)arr[i]=bin.charCodeAt(i);
const url=URL.createObjectURL(new Blob([arr],{type:"application/pdf"}));
document.getElementById("o").onclick=()=>window.open(url,"_blank");
document.getElementById("p").onclick=()=>{
  document.querySelectorAll("iframe").forEach(f=>f.remove());
  const f=document.createElement("iframe");
  f.style.cssText="position:fixed;right:0;bottom:0;width:0;height:0;border:0";
  f.onload=()=>{try{f.contentWindow.focus();f.contentWindow.print();}catch(e){window.open(url,"_blank");}};
  f.src=url;document.body.appendChild(f);
};
</script>"""


def print_bar(pdf):
    """Tombol cetak langsung (dialog cetak browser) + buka PDF di tab baru."""
    html = PRINT_HTML.replace("__B64__", base64.b64encode(pdf).decode())
    if hasattr(st, "iframe"):
        st.iframe(html, height=46)
    else:
        import streamlit.components.v1 as comp
        comp.html(html, height=46)


@st.cache_data(show_spinner=False, max_entries=30)
def render_pages(pdf, dpi=120):
    if pymupdf is None:
        return []
    with pymupdf.open(stream=pdf, filetype="pdf") as doc:
        return [p.get_pixmap(dpi=dpi).tobytes("png") for p in doc]


def preview(pdf, mode):
    imgs = render_pages(pdf)
    if not imgs:
        st.info("Pratinjau butuh paket `pymupdf` (lihat requirements.txt).")
        return
    forms = len(imgs) - (1 if mode == "semua" else 0)
    labels = [] if mode == "amplop" else (["Tanda terima"] if forms == 1 else
                                          [f"Tanda terima {i}/{forms}" for i in range(1, forms + 1)])
    labels += ["Amplop"] if mode in ("amplop", "semua") else []
    if len(imgs) == 1:
        st.image(imgs[0], width=820)
    else:
        for tab, img in zip(st.tabs(labels), imgs):
            tab.image(img, width=820)


def show_receipt(rid, key):
    opts = {"Tanda terima": "tt", "Amplop": "amplop", "Keduanya": "semua"}
    pick = st.segmented_control("Yang dicetak / diunduh", list(opts), default="Tanda terima",
                                key=f"mode_{key}") or "Tanda terima"
    mode = opts[pick]
    pdf = pdf_for(rid, mode)
    d = D("receipts")
    no = d.loc[d.id == rid, "nomor"].iloc[0]
    a, b = st.columns([3, 2])
    with a:
        print_bar(pdf)
    b.download_button("⬇️ Unduh PDF", pdf, f"{no}.pdf", "application/pdf", key=f"dl_{key}")
    st.caption("Di dialog cetak: ukuran kertas **A5** (atau A4 dengan skala **100% / Actual size**, lalu potong "
               "setengah), margin **None/Tanpa**. Jika dialog tidak muncul, pakai **Buka di tab baru** lalu Ctrl+P.")
    preview(pdf, mode)


def next_nomor(jenis, tahun):
    d = D("receipts")
    m = d[(d.jenis == jenis) & (d.tahun == tahun)].seq.max()
    return f"{(int(m) + 1 if pd.notna(m) else 1):04d}-{jenis}"


# ---------- Halaman ----------
def page_buat():
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
    cl = D("clients").sort_values("nama")
    cr = D("couriers")
    cr = cr[cr.aktif == 1]
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
    kontak = [x.strip() for x in str(row.kontak or "").split("\n") if x.strip()]
    if kontak:
        up = st.selectbox("U.p. (pilih atau ketik baru)", kontak, index=0, accept_new_options=True)
    else:
        up = st.text_input("U.p. (opsional)")
    st.write("**Daftar dokumen**")
    items = st.data_editor(pd.DataFrame([{"Uraian": "", "Jumlah": "1 asli", "Keterangan": ""}]),
                           num_rows="dynamic", width="stretch", key=f"items{n}")
    catatan = st.text_input("Catatan (opsional)")
    its = [r for r in items.fillna("").to_dict("records") if str(r["Uraian"]).strip()]

    st.subheader("Pratinjau")
    rprev = dict(nomor=next_nomor(jenis, tgl.year), tanggal=tgl.isoformat(), jam=jam.strftime("%H:%M"),
                 up=up or "", catatan=catatan)
    preview(make_pdf(rprev, row, cr.set_index("id").nama[kur], get_kop(), its, "tt"), "tt")
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
    st.dataframe(v[["nomor", "jenis", "tanggal", "jam", "klien", "pengantar", "status"]],
                 hide_index=True, width="stretch")
    if v.empty:
        return
    st.divider()
    lab = {r.id: f"{r.nomor} · {r.tanggal} · {r.klien}" for r in v.itertuples()}
    rid = st.selectbox("Pilih tanda terima", list(lab), format_func=lab.get)
    r = v[v.id == rid].iloc[0]
    a, b = st.columns(2)
    ns = a.selectbox("Status", STATUS, index=STATUS.index(r.status) if r.status in STATUS else 0)
    if a.button("Perbarui status"):
        W("update", "receipts", int(rid), {"status": ns})
        st.rerun()
    if b.checkbox("Izinkan hapus") and b.button("🗑️ Hapus"):
        W("delete", "receipts", int(rid))
        st.rerun()
    show_receipt(int(rid), f"h{rid}")


def page_dasbor():
    st.header("Dasbor")
    df = load_receipts()
    if df.empty:
        st.info("Belum ada data.")
        return
    hari = st.number_input("Dianggap 'belum kembali' setelah (hari)", 1, 90, 7)
    cutoff = (datetime.now(WIB).date() - timedelta(days=hari)).isoformat()
    over = df[(df.status == "Diserahkan") & (df.tanggal < cutoff)]
    bulan = datetime.now(WIB).strftime("%Y-%m")
    m1, m2, m3 = st.columns(3)
    m1.metric("Total tanda terima", len(df))
    m2.metric("Bulan ini", int((df.tanggal.str[:7] == bulan).sum()))
    m3.metric("Belum kembali", len(over))
    st.subheader("Per bulan")
    st.bar_chart(df.assign(bulan=df.tanggal.str[:7]).groupby(["bulan", "jenis"]).size().unstack(fill_value=0))
    st.subheader("Per pengantar")
    st.bar_chart(df.pengantar.fillna("-").value_counts())
    if not over.empty:
        st.subheader("⚠️ Belum kembali")
        st.dataframe(over[["nomor", "tanggal", "klien", "pengantar"]], hide_index=True, width="stretch")


def page_master():
    st.header("Klien & Pengantar")
    st.subheader("Klien")
    st.caption("Kolom *kontak*: satu nama (U.p.) per baris. Hapus baris dengan memilihnya lalu tekan Delete.")
    ed = st.data_editor(D("clients").sort_values("nama"), num_rows="dynamic", width="stretch", hide_index=True,
                        key="ce", disabled=["id"],
                        column_config={"kontak": st.column_config.TextColumn(width="large")})
    if st.button("Simpan klien"):
        sync("clients", ed, ["nama", "kode", "alamat1", "alamat2", "alamat3", "alamat4", "kontak"],
             "client_id", "Klien")
        st.rerun()
    st.subheader("Pengantar")
    kdf = D("couriers").sort_values("nama")
    kdf["aktif"] = kdf["aktif"].astype(bool)
    ek = st.data_editor(kdf, num_rows="dynamic", width="stretch", hide_index=True, key="ke",
                        disabled=["id"], column_config={"aktif": st.column_config.CheckboxColumn()})
    if st.button("Simpan pengantar"):
        sync("couriers", ek.fillna({"aktif": True}), ["nama", "aktif"], "courier_id", "Pengantar")
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
        ada_klien = set(D("clients", True).nama)
        if "PT CONTOH ABADI Tbk" not in ada_klien:
            W("insert", "clients", dict(nama="PT CONTOH ABADI Tbk", kode="CABD", alamat1="Menara Contoh Lt. 10",
                                        alamat2="Jl. Jenderal Sudirman Kav. 1", alamat3="Jakarta 10220", alamat4="",
                                        kontak="Bpk. Budi\nIbu Sari"))
        ada_kurir = set(D("couriers", True).nama)
        for n in ("Pengantar A", "Pengantar B"):
            if n not in ada_kurir:
                W("insert", "couriers", {"nama": n, "aktif": 1})
        st.success("Data contoh ditambahkan bila belum ada.")


# ---------- Main ----------
pw = secret("app_password")
if pw and not st.session_state.get("ok"):
    st.title("Tanda Terima")
    sandi = st.text_input("Kata sandi", type="password")
    if sandi and sandi != pw:
        st.error("Kata sandi salah.")
    if sandi == pw:
        st.session_state["ok"] = True
        st.rerun()
    st.stop()

PAGES = {"🧾 Buat": page_buat, "📋 Riwayat": page_riwayat, "📊 Dasbor": page_dasbor,
         "👥 Klien & Pengantar": page_master, "⚙️ Pengaturan": page_setelan}
st.sidebar.title("Tanda Terima")
PAGES[st.sidebar.radio("Menu", list(PAGES), label_visibility="collapsed")]()
