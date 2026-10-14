"""Tanda Terima — prototipe pengganti OmniForm (Streamlit + Google Sheets / SQLite + ReportLab)."""
import io, json, sqlite3
from datetime import datetime, timedelta
from xml.sax.saxutils import escape
from zoneinfo import ZoneInfo

import pandas as pd
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfgen import canvas
from reportlab.platypus import Paragraph, Table, TableStyle

st.set_page_config(page_title="Tanda Terima", page_icon="🧾", layout="wide")
WIB = ZoneInfo("Asia/Jakarta")
JENIS = {"PR": "Proposal", "INV": "Invoice", "KW": "Kwitansi", "LP": "Laporan"}
STATUS = ["Draft", "Diserahkan", "Diterima", "Arsip"]
BLN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
       "September", "Oktober", "November", "Desember"]
KOP = {"nama": "SUWENDHO RINALDY & REKAN",
       "kiri": "KANTOR JASA PENILAI PUBLIK\nNomor Izin Usaha KJPP: 2.09.0059\nNomor Izin Cabang KJPP: 1138/KM.1/2017\nPenilai Properti dan Bisnis",
       "kanan": "Kantor Cabang Jakarta\nKomplek Kalibata Indah Blok K16-17\nJl. Rawajati Timur, Pancoran\nJakarta Selatan 12750\nT (021) 7970913 / 799-4521"}
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
            if w.row_values(1) != cols:
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


# ---------- PDF ----------
def make_pdf(r, cl, kurir, kop, items):
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    W, H = A4
    M = 18 * mm
    navy = colors.HexColor("#0b1f6b")
    c.setFillColor(navy)
    c.setFont("Helvetica-Bold", 13)
    c.drawString(M, H - 16 * mm, kop["nama"])
    c.setFont("Helvetica", 8)
    y = H - 20.5 * mm
    for ln in kop["kiri"].split("\n"):
        c.drawString(M, y, ln)
        y -= 3.8 * mm
    y = H - 16 * mm
    for n, ln in enumerate(kop["kanan"].split("\n")):
        c.setFont("Helvetica-Bold" if n == 0 else "Helvetica", 8)
        c.drawRightString(W - M, y, ln)
        y -= 3.8 * mm
    c.setStrokeColor(navy)
    c.setLineWidth(1.2)
    c.line(M, H - 42 * mm, W - M, H - 42 * mm)

    c.setFillColor(colors.black)
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(W / 2, H - 53 * mm, "TANDA TERIMA")
    c.setFont("Helvetica-Oblique", 9)
    c.drawCentredString(W / 2, H - 58 * mm, "R E C E I P T")

    top = H - 68 * mm
    c.setFont("Helvetica", 8)
    c.drawString(M, top, "Kepada / To")
    c.setFont("Helvetica-Bold", 10)
    y = top - 5 * mm
    c.drawString(M, y, cl["nama"])
    c.setFont("Helvetica", 9)
    for k in ("alamat1", "alamat2", "alamat3", "alamat4"):
        if cl[k]:
            y -= 4.2 * mm
            c.drawString(M, y, cl[k])
    if r["up"]:
        y -= 5 * mm
        c.drawString(M, y, f"U.p. : {r['up']}")
    x = W - M - 62 * mm
    for i, (lab, val) in enumerate([("No. / Number", r["nomor"]), ("Tgl / Date", idn(r["tanggal"])),
                                    ("Jam / Hour", r["jam"])]):
        c.setFont("Helvetica", 8)
        c.drawString(x, top - i * 5.5 * mm, lab)
        c.setFont("Helvetica-Bold", 9)
        c.drawString(x + 22 * mm, top - i * 5.5 * mm, f": {val}")

    y = min(y, top - 20 * mm) - 9 * mm
    c.setFont("Helvetica", 9)
    c.drawString(M, y, "Telah diterima/diserahkan  (Received/delivered):")
    P = ParagraphStyle("p", fontName="Helvetica", fontSize=9, leading=11)
    B = ParagraphStyle("b", parent=P, fontName="Helvetica-Bold")
    rows = [[Paragraph("No.", B), Paragraph("Uraian / Description", B), Paragraph("Jumlah / Qty", B),
             Paragraph("Keterangan / Remarks", B)]]
    for n, it in enumerate(items, 1):
        rows.append([Paragraph(f"{n}.", P)] + [Paragraph(escape(str(it.get(k) or "")), P)
                                               for k in ("Uraian", "Jumlah", "Keterangan")])
    t = Table(rows, colWidths=[10 * mm, 78 * mm, 24 * mm, W - 2 * M - 112 * mm])
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black), ("VALIGN", (0, 0), (-1, -1), "TOP"),
                           ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#e8ebf5"))]))
    _, th = t.wrapOn(c, W - 2 * M, H)
    t.drawOn(c, M, y - 4 * mm - th)
    y = y - 4 * mm - th - 12 * mm
    if y < 65 * mm:
        c.showPage()
        y = H - 40 * mm
    c.setFont("Helvetica", 9)
    c.drawString(M, y, "Diterima oleh / Received by")
    c.drawString(x, y, "Diserahkan oleh / Delivered by")
    c.line(M, y - 24 * mm, M + 55 * mm, y - 24 * mm)
    c.line(x, y - 24 * mm, x + 55 * mm, y - 24 * mm)
    c.setFont("Helvetica", 8)
    c.drawString(M, y - 28 * mm, "Tgl/Jam: ........................")
    c.drawString(x, y - 28 * mm, kurir or "")
    if r["catatan"]:
        c.drawString(M, y - 40 * mm, f"Catatan: {r['catatan']}")

    c.showPage()  # halaman 2: amplop
    c.setFont("Helvetica", 12)
    c.drawString(M + 20 * mm, H - 60 * mm, "Kepada Yth.")
    c.setFont("Helvetica-Bold", 15)
    y = H - 70 * mm
    c.drawString(M + 20 * mm, y, cl["nama"])
    c.setFont("Helvetica", 12)
    for k in ("alamat1", "alamat2", "alamat3", "alamat4"):
        if cl[k]:
            y -= 6.5 * mm
            c.drawString(M + 20 * mm, y, cl[k])
    if r["up"]:
        c.setFont("Helvetica-Bold", 12)
        c.drawString(M + 20 * mm, y - 12 * mm, f"U.p. : {r['up']}")
    c.save()
    return buf.getvalue()


def pdf_for(rid):
    d = D("receipts")
    r = d[d.id == rid].iloc[0]
    cl = D("clients")
    cl = cl[cl.id == r.client_id].iloc[0]
    k = D("couriers")
    k = k[k.id == r.courier_id]
    return make_pdf(r, cl, k.nama.iloc[0] if len(k) else "", get_kop(), json.loads(r["items"]))


# ---------- Halaman ----------
def page_buat():
    st.header("Buat tanda terima")
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
    kontak = [x.strip() for x in row.kontak.split("\n") if x.strip()]
    up = st.selectbox("U.p. (pilih atau ketik baru)", kontak, index=0 if kontak else None,
                      accept_new_options=True)
    st.write("**Daftar dokumen**")
    items = st.data_editor(pd.DataFrame([{"Uraian": "", "Jumlah": "1 asli", "Keterangan": ""}]),
                           num_rows="dynamic", width="stretch", key=f"items{n}")
    catatan = st.text_input("Catatan (opsional)")
    if st.button("💾 Simpan & buat PDF", type="primary"):
        its = [r for r in items.fillna("").to_dict("records") if str(r["Uraian"]).strip()]
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
    if "last" in st.session_state:
        rid = st.session_state["last"]
        d = D("receipts")
        no = d.loc[d.id == rid, "nomor"].iloc[0]
        st.success(f"Tanda terima **{no}** tersimpan.")
        st.download_button("⬇️ Unduh PDF (tanda terima + amplop)", pdf_for(rid), f"{no}.pdf", "application/pdf")


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
    st.table(pd.DataFrame(json.loads(r["items"])))
    a, b, c3 = st.columns(3)
    ns = a.selectbox("Status", STATUS, index=STATUS.index(r.status) if r.status in STATUS else 0)
    if a.button("Perbarui status"):
        W("update", "receipts", int(rid), {"status": ns})
        st.rerun()
    b.download_button("⬇️ PDF", pdf_for(int(rid)), f"{r.nomor}.pdf", "application/pdf")
    if c3.checkbox("Izinkan hapus") and c3.button("🗑️ Hapus"):
        W("delete", "receipts", int(rid))
        st.rerun()


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

PAGES = {"🧾 Buat": page_buat, "📋 Riwayat": page_riwayat, "📊 Dasbor": page_dasbor,
         "👥 Klien & Pengantar": page_master, "⚙️ Pengaturan": page_setelan}
st.sidebar.title("Tanda Terima")
PAGES[st.sidebar.radio("Menu", list(PAGES), label_visibility="collapsed")]()
