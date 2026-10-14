"""Tanda Terima — prototipe pengganti OmniForm (Streamlit + SQLAlchemy + ReportLab)."""
import io, json
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
from sqlalchemy import (Column, Integer, MetaData, String, Table as T, Text,
                        UniqueConstraint, create_engine, delete, func, insert, select, text, update)
from sqlalchemy.exc import IntegrityError

st.set_page_config(page_title="Tanda Terima", page_icon="🧾", layout="wide")
WIB = ZoneInfo("Asia/Jakarta")
JENIS = {"PR": "Proposal", "INV": "Invoice", "KW": "Kwitansi", "LP": "Laporan"}
STATUS = ["Draft", "Diserahkan", "Diterima", "Arsip"]
BLN = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus",
       "September", "Oktober", "November", "Desember"]
KOP = {"nama": "SUWENDHO RINALDY & REKAN",
       "kiri": "KANTOR JASA PENILAI PUBLIK\nNomor Izin Usaha KJPP: 2.09.0059\nNomor Izin Cabang KJPP: 1138/KM.1/2017\nPenilai Properti dan Bisnis",
       "kanan": "Kantor Cabang Jakarta\nKomplek Kalibata Indah Blok K16-17\nJl. Rawajati Timur, Pancoran\nJakarta Selatan 12750\nT (021) 7970913 / 799-4521"}


def secret(k, d=None):
    try:
        return st.secrets.get(k, d)
    except Exception:
        return d


# ---------- Database ----------
@st.cache_resource
def get_engine():
    url = str(secret("database_url", "sqlite:///tandaterima.db")).strip()
    for pre in ("postgres://", "postgresql://", "postgresql+psycopg2://"):  # pakai driver psycopg v3
        if url.startswith(pre):
            url = "postgresql+psycopg://" + url[len(pre):]
    kw = {"connect_args": {"connect_timeout": 10}} if url.startswith("postgresql") else {}
    return create_engine(url, pool_pre_ping=True, **kw)


eng = get_engine()
md = MetaData()
clients = T("clients", md, Column("id", Integer, primary_key=True, autoincrement=True),
            Column("nama", String(200)), Column("kode", String(20)), Column("alamat1", String(200)),
            Column("alamat2", String(200)), Column("alamat3", String(200)), Column("alamat4", String(200)),
            Column("kontak", Text))
couriers = T("couriers", md, Column("id", Integer, primary_key=True, autoincrement=True),
             Column("nama", String(100)), Column("aktif", Integer, default=1))
receipts = T("receipts", md, Column("id", Integer, primary_key=True, autoincrement=True),
             Column("jenis", String(5)), Column("tahun", Integer), Column("seq", Integer),
             Column("nomor", String(30)), Column("tanggal", String(10)), Column("jam", String(5)),
             Column("client_id", Integer), Column("up", String(200)), Column("courier_id", Integer),
             Column("status", String(20)), Column("catatan", Text), Column("items", Text),
             Column("created_at", String(25)), UniqueConstraint("jenis", "tahun", "seq"))
settings = T("settings", md, Column("k", String(50), primary_key=True), Column("v", Text))
try:
    with eng.connect() as _c:
        _c.execute(text("select 1"))
except Exception as _e:  # tampilkan diagnosis (pesan asli Streamlit disembunyikan)
    _u, _o = eng.url, getattr(_e, "orig", None) or _e
    st.error("Tidak bisa terhubung ke database.")
    st.code(f"host={_u.host}  port={_u.port}  user={_u.username}  db={_u.database}\n"
            f"{type(_o).__name__}: {_o}")
    st.markdown(
        "**Cek berikut:**\n"
        "1. **Supabase:** pakai URL *Connect → Session pooler* (host `aws-...pooler.supabase.com`, user "
        "`postgres.<project-ref>`). URL *direct* (`db.<ref>.supabase.co`) hanya IPv6 dan gagal dari Streamlit Cloud.\n"
        "2. Hapus tanda kurung siku `[ ]` di sekitar password pada URL contoh.\n"
        "3. Password dengan karakter khusus harus di-encode (`@`→`%40`, `/`→`%2F`, `#`→`%23`, `:`→`%3A`).\n"
        "4. **Neon:** tambahkan `?sslmode=require` di akhir URL.\n"
        "5. Ingin mencoba dulu tanpa database tetap? Hapus baris `database_url` di Secrets (memakai SQLite sementara).")
    st.stop()
md.create_all(eng)


def q(stmt):
    with eng.connect() as c:
        return pd.read_sql(stmt, c)


def run(stmt):
    with eng.begin() as c:
        c.execute(stmt)


def get_kop():
    kop = dict(KOP)
    kop.update({r.k: r.v for r in q(select(settings)).itertuples()})
    return kop


def idn(iso):
    d = datetime.fromisoformat(iso)
    return f"{d.day} {BLN[d.month - 1]} {d.year}"


def save_receipt(data):
    for _ in range(3):  # ulangi bila bentrok nomor (dua pengguna bersamaan)
        with eng.connect() as c:
            m = c.execute(select(func.max(receipts.c.seq)).where(
                receipts.c.jenis == data["jenis"], receipts.c.tahun == data["tahun"])).scalar()
        seq = (m or 0) + 1
        try:
            with eng.begin() as c:
                r = c.execute(insert(receipts).values(**data, seq=seq, nomor=f"{seq:04d}-{data['jenis']}"))
                return r.inserted_primary_key[0]
        except IntegrityError:
            continue
    raise RuntimeError("Gagal membuat nomor unik, coba lagi.")


def sync(tbl, df, cols, ref_col, label):
    """Simpan hasil st.data_editor: tambah / ubah / hapus (hapus diblokir bila sudah dipakai)."""
    old = q(select(tbl.c.id))["id"].tolist()
    keep = set()
    with eng.begin() as c:
        for _, r in df.iterrows():
            vals = {}
            for k in cols:
                v = r[k]
                v = None if pd.isna(v) else (v.item() if hasattr(v, "item") else v)
                vals[k] = int(v) if isinstance(v, bool) else v
            if not str(vals[cols[0]] or "").strip():
                continue
            if pd.isna(r["id"]):
                c.execute(insert(tbl).values(**vals))
            else:
                c.execute(update(tbl).where(tbl.c.id == int(r["id"])).values(**vals))
                keep.add(int(r["id"]))
        for i in old:
            if i in keep:
                continue
            used = c.execute(select(func.count()).select_from(receipts).where(ref_col == i)).scalar()
            if used:
                st.warning(f"{label} id {i} sudah dipakai di tanda terima, tidak dihapus (nonaktifkan saja).")
            else:
                c.execute(delete(tbl).where(tbl.c.id == i))


def load_receipts():
    r = q(select(receipts))
    c = q(select(clients.c.id.label("client_id"), clients.c.nama.label("klien")))
    k = q(select(couriers.c.id.label("courier_id"), couriers.c.nama.label("pengantar")))
    return r.merge(c, on="client_id", how="left").merge(k, on="courier_id", how="left")


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
    r = q(select(receipts).where(receipts.c.id == rid)).iloc[0]
    cl = q(select(clients).where(clients.c.id == int(r.client_id))).iloc[0].fillna("")
    k = q(select(couriers.c.nama).where(couriers.c.id == int(r.courier_id)))
    return make_pdf(r.fillna(""), cl, k.iloc[0, 0] if len(k) else "", get_kop(), json.loads(r["items"]))


# ---------- Halaman ----------
def page_buat():
    st.header("Buat tanda terima")
    cl = q(select(clients).order_by(clients.c.nama))
    cr = q(select(couriers).where(couriers.c.aktif == 1))
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
    kontak = [x.strip() for x in (row.kontak or "").split("\n") if x.strip()]
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
        no = q(select(receipts.c.nomor).where(receipts.c.id == rid)).iloc[0, 0]
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
        hay = (v.klien.fillna("") + v.nomor + v["items"].fillna("")).str.lower()
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
    ns = a.selectbox("Status", STATUS, index=STATUS.index(r.status))
    if a.button("Perbarui status"):
        run(update(receipts).where(receipts.c.id == int(rid)).values(status=ns))
        st.rerun()
    b.download_button("⬇️ PDF", pdf_for(int(rid)), f"{r.nomor}.pdf", "application/pdf")
    if c3.checkbox("Izinkan hapus") and c3.button("🗑️ Hapus"):
        run(delete(receipts).where(receipts.c.id == int(rid)))
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
    cdf = q(select(clients).order_by(clients.c.nama))
    ed = st.data_editor(cdf, num_rows="dynamic", width="stretch", hide_index=True, key="ce",
                        disabled=["id"], column_config={"kontak": st.column_config.TextColumn(width="large")})
    if st.button("Simpan klien"):
        sync(clients, ed, ["nama", "kode", "alamat1", "alamat2", "alamat3", "alamat4", "kontak"],
             receipts.c.client_id, "Klien")
        st.rerun()
    st.subheader("Pengantar")
    kdf = q(select(couriers).order_by(couriers.c.nama))
    kdf["aktif"] = kdf["aktif"].astype(bool)
    ek = st.data_editor(kdf, num_rows="dynamic", width="stretch", hide_index=True, key="ke",
                        disabled=["id"], column_config={"aktif": st.column_config.CheckboxColumn()})
    if st.button("Simpan pengantar"):
        sync(couriers, ek.fillna({"aktif": True}), ["nama", "aktif"], receipts.c.courier_id, "Pengantar")
        st.rerun()


def page_setelan():
    st.header("Pengaturan")
    st.subheader("Kop surat")
    kop = get_kop()
    nama = st.text_input("Nama KJPP", kop["nama"])
    kiri = st.text_area("Blok kiri (di bawah nama)", kop["kiri"], height=110)
    kanan = st.text_area("Blok kanan (alamat cabang)", kop["kanan"], height=130)
    if st.button("Simpan kop"):
        for k, v in dict(nama=nama, kiri=kiri, kanan=kanan).items():
            run(delete(settings).where(settings.c.k == k))
            run(insert(settings).values(k=k, v=v))
        st.success("Tersimpan.")
    st.subheader("Cadangan data")
    st.warning("Di Streamlit Cloud, file SQLite **hilang saat aplikasi restart/redeploy**. Untuk data tetap, isi "
               "`database_url` (Postgres, mis. Supabase/Neon) di *Secrets*, atau unduh cadangan secara rutin.")
    data = {t.name: q(select(t)).astype(object).where(lambda d: d.notna(), None).to_dict("records")
            for t in (clients, couriers, receipts, settings)}
    st.download_button("⬇️ Unduh cadangan (JSON)", json.dumps(data, ensure_ascii=False, default=str, indent=1),
                       f"cadangan-tanda-terima-{datetime.now(WIB):%Y%m%d}.json", "application/json")
    up = st.file_uploader("Pulihkan dari cadangan (menimpa semua data)", type="json")
    if up and st.button("Pulihkan sekarang"):
        d = json.load(up)
        with eng.begin() as c:
            for t in (receipts, clients, couriers, settings):
                c.execute(delete(t))
                if d.get(t.name):
                    c.execute(insert(t), d[t.name])
        st.success("Data dipulihkan.")
    st.subheader("Data contoh")
    if st.button("Isi data contoh (fiktif)"):
        run(insert(clients).values(nama="PT CONTOH ABADI Tbk", kode="CABD", alamat1="Menara Contoh Lt. 10",
                                   alamat2="Jl. Jenderal Sudirman Kav. 1", alamat3="Jakarta 10220", alamat4="",
                                   kontak="Bpk. Budi\nIbu Sari"))
        with eng.begin() as c:
            c.execute(insert(couriers), [{"nama": "Pengantar A", "aktif": 1}, {"nama": "Pengantar B", "aktif": 1}])
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
