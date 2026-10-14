import { useEffect, useMemo, useState } from "react";
import {
  AlertTriangle,
  Download,
  FilePlus2,
  LayoutDashboard,
  Plus,
  Printer,
  ScrollText,
  Settings,
  Trash2,
  Users,
} from "lucide-react";
import { Bar, BarChart, ResponsiveContainer, Tooltip, XAxis, YAxis } from "recharts";
import {
  type Client,
  type Courier,
  type Item,
  type Jenis,
  type Receipt,
  type Status,
  EMPTY_CLIENT,
  JENIS,
  JENIS_KEYS,
  KOP,
  STATUS,
  addDays,
  blankItem,
  contactsOf,
  duplicateNomors,
  formatNomor,
  nextSeq,
  nowJakarta,
} from "@/lib/tanda/model";
import { parseBackup, toBackup, useDesk, type DeskData } from "@/lib/tanda/store";
import { SlipPreview, type PrintMode } from "@/components/tanda/sheet";

type PageId = "buat" | "riwayat" | "dasbor" | "klien" | "atur";

const NAV: { id: PageId; label: string; icon: typeof FilePlus2 }[] = [
  { id: "buat", label: "Buat", icon: FilePlus2 },
  { id: "riwayat", label: "Riwayat", icon: ScrollText },
  { id: "dasbor", label: "Dasbor", icon: LayoutDashboard },
  { id: "klien", label: "Klien", icon: Users },
  { id: "atur", label: "Pengaturan", icon: Settings },
];

const MODES: { id: PrintMode; label: string }[] = [
  { id: "tt", label: "Tanda terima" },
  { id: "amplop", label: "Amplop" },
  { id: "semua", label: "Keduanya" },
];

function Mark() {
  return (
    <span className="grid h-9 w-9 shrink-0 place-items-center rounded-md bg-navy text-[11px] font-bold leading-none tracking-tight text-form-paper">
      SR
    </span>
  );
}

function ModePicker({ mode, onChange }: { mode: PrintMode; onChange: (m: PrintMode) => void }) {
  return (
    <div role="radiogroup" aria-label="Yang dicetak" className="no-print flex flex-wrap gap-2">
      {MODES.map((m) => (
        <button
          key={m.id}
          type="button"
          role="radio"
          aria-checked={mode === m.id}
          className="btn btn-line"
          onClick={() => onChange(m.id)}
        >
          {m.label}
        </button>
      ))}
    </div>
  );
}

function openPrintTab(id: number, mode: PrintMode) {
  const url = new URL(window.location.href);
  url.searchParams.set("cetak", String(id));
  url.searchParams.set("mode", mode);
  window.open(url.toString(), "_blank", "noopener");
}

function homeHref() {
  return window.location.pathname;
}

function PrintActions({ id, mode }: { id: number; mode: PrintMode }) {
  return (
    <div className="no-print flex flex-wrap items-center gap-3">
      <button type="button" className="btn btn-primary" onClick={() => window.print()}>
        <Printer className="size-4" aria-hidden />
        Cetak
      </button>
      <button type="button" className="btn btn-line" onClick={() => openPrintTab(id, mode)}>
        Buka di tab baru
      </button>
      <p className="max-w-md text-xs leading-relaxed text-muted">
        Di dialog cetak pilih kertas A5 dan margin none. Bisa juga A4 pada skala 100%, lalu potong setengah.
        Untuk berkas PDF, pilih Simpan sebagai PDF. Jika dialog tidak muncul, buka di tab baru lalu cetak dari sana.
      </p>
    </div>
  );
}

function clientName(clients: Client[], id: number): string {
  return clients.find((c) => c.id === id)?.nama ?? "Klien tidak ditemukan";
}

function courierName(couriers: Courier[], id: number): string {
  return couriers.find((c) => c.id === id)?.nama ?? "";
}

const CHART: { key: Jenis; fill: string; opacity: number }[] = [
  { key: "PR", fill: "var(--color-navy)", opacity: 1 },
  { key: "INV", fill: "var(--color-kop)", opacity: 1 },
  { key: "KW", fill: "var(--color-muted)", opacity: 1 },
  { key: "LP", fill: "var(--color-navy)", opacity: 0.45 },
];

function CetakPage({ id, mode }: { id: number; mode: PrintMode }) {
  const receipt = useDesk((s) => s.receipts.find((r) => r.id === id) ?? null);
  const clients = useDesk((s) => s.clients);
  const couriers = useDesk((s) => s.couriers);
  const kop = useDesk((s) => s.kop);

  if (!receipt) {
    return (
      <div className="no-print mx-auto max-w-lg px-4 py-8">
        <p className="text-sm">Tanda terima ini tidak ada di peramban ini.</p>
        <a className="btn btn-line mt-4" href={homeHref()}>
          Kembali
        </a>
      </div>
    );
  }

  return (
    <div className="desk-shell bg-paper">
      <div className="no-print mx-auto flex max-w-3xl flex-wrap items-center gap-3 px-4 py-4">
        <button type="button" className="btn btn-primary" onClick={() => window.print()}>
          <Printer className="size-4" aria-hidden />
          Cetak
        </button>
        <a className="btn btn-line" href={homeHref()}>
          Kembali
        </a>
        <p className="max-w-md text-xs leading-relaxed text-muted">
          Kertas A5, margin none. Untuk PDF, pilih Simpan sebagai PDF. Nomor {receipt.nomor}.
        </p>
      </div>
      <SlipPreview
        mode={mode}
        slip={{
          nomor: receipt.nomor,
          tanggal: receipt.tanggal,
          jam: receipt.jam,
          up: receipt.up,
          catatan: receipt.catatan,
          items: receipt.items,
          client: clients.find((c) => c.id === receipt.client_id) ?? null,
          kurir: courierName(couriers, receipt.courier_id),
          kop,
        }}
      />
    </div>
  );
}

export function Desk() {
  const [page, setPage] = useState<PageId>("buat");
  const [job, setJob] = useState<{ id: number; mode: PrintMode } | null>(null);
  useEffect(() => {
    const q = new URLSearchParams(window.location.search);
    const id = Number(q.get("cetak"));
    const mode = q.get("mode");
    const isPrint =
      Number.isFinite(id) && id > 0 && (mode === "tt" || mode === "amplop" || mode === "semua");
    void Promise.resolve(useDesk.persist.rehydrate()).finally(() => {
      if (isPrint) setJob({ id, mode: mode as PrintMode });
    });
  }, []);

  if (job) return <CetakPage id={job.id} mode={job.mode} />;

  return (
    <div className="desk-shell min-h-dvh lg:grid lg:grid-cols-[15rem_minmax(0,1fr)]">
      <aside className="no-print hidden border-r border-line bg-sheet lg:flex lg:flex-col">
        <div className="flex items-center gap-3 px-4 py-5">
          <Mark />
          <div>
            <p className="text-sm font-semibold leading-tight">Tanda Terima</p>
            <p className="text-xs text-muted">KJPP SRR</p>
          </div>
        </div>
        <nav className="flex flex-col gap-1 px-3" aria-label="Menu">
          {NAV.map((item) => (
            <button
              key={item.id}
              type="button"
              aria-current={page === item.id ? "page" : undefined}
              className="btn btn-line justify-start"
              onClick={() => setPage(item.id)}
            >
              <item.icon className="size-4" aria-hidden />
              {item.label}
            </button>
          ))}
        </nav>
        <p className="mt-auto px-4 py-5 text-xs leading-relaxed text-muted">
          Tata letak formulir mengikuti template OmniForm dan tidak diubah. Data tersimpan di peramban ini.
        </p>
      </aside>

      <div className="min-w-0">
        <header className="no-print sticky top-0 z-20 border-b border-line bg-paper/95 backdrop-blur lg:hidden">
          <div className="flex items-center gap-3 px-4 py-3">
            <Mark />
            <p className="text-sm font-semibold">Tanda Terima</p>
          </div>
          <nav className="flex gap-2 overflow-x-auto px-3 pb-3" aria-label="Menu">
            {NAV.map((item) => (
              <button
                key={item.id}
                type="button"
                aria-current={page === item.id ? "page" : undefined}
                className="btn btn-line shrink-0"
                onClick={() => setPage(item.id)}
              >
                {item.label}
              </button>
            ))}
          </nav>
        </header>
        <main className="mx-auto w-full max-w-6xl px-4 py-5 sm:px-6">
          {page === "buat" ? <Buat /> : null}
          {page === "riwayat" ? <Riwayat /> : null}
          {page === "dasbor" ? <Dasbor /> : null}
          {page === "klien" ? <KlienPage /> : null}
          {page === "atur" ? <Atur /> : null}
        </main>
      </div>
    </div>
  );
}

function Buat() {
  const clients = useDesk((s) => s.clients);
  const couriers = useDesk((s) => s.couriers);
  const receipts = useDesk((s) => s.receipts);
  const kop = useDesk((s) => s.kop);
  const saveReceipt = useDesk((s) => s.saveReceipt);
  const addSamples = useDesk((s) => s.addSamples);

  const sorted = useMemo(
    () => [...clients].sort((a, b) => a.nama.localeCompare(b.nama, "id")),
    [clients],
  );
  const active = couriers.filter((c) => c.aktif);

  const [jenis, setJenis] = useState<Jenis>("PR");
  const [clock, setClock] = useState({ date: "2026-10-14", time: "09:00" });
  const [clientId, setClientId] = useState(sorted[0]?.id ?? 0);
  const [courierId, setCourierId] = useState(active[0]?.id ?? 0);
  const [up, setUp] = useState(contactsOf(sorted[0])[0] ?? "");
  const [catatan, setCatatan] = useState("");
  const [rows, setRows] = useState<{ key: string; item: Item }[]>([{ key: "awal", item: blankItem() }]);
  const [error, setError] = useState("");
  const [last, setLast] = useState<Receipt | null>(null);
  const [mode, setMode] = useState<PrintMode>("tt");

  useEffect(() => {
    const n = nowJakarta();
    setClock({ date: n.date, time: n.time });
  }, []);

  const client = sorted.find((c) => c.id === clientId) ?? sorted[0] ?? null;
  const courier = active.find((c) => c.id === courierId) ?? active[0] ?? null;
  const tahun = Number(clock.date.slice(0, 4)) || 2026;
  const previewNo = formatNomor(nextSeq(receipts, jenis, tahun), jenis);
  const ups = contactsOf(client ?? undefined);

  if (!client || !courier) {
    return (
      <section className="panel max-w-xl p-5">
        <h1 className="text-xl font-semibold">Buat tanda terima</h1>
        <p className="mt-2 text-sm leading-relaxed text-muted">
          Isi dulu klien dan pengantar yang masih aktif. Data contoh bersifat fiktif dan bisa dihapus.
        </p>
        <div className="mt-4 flex flex-wrap gap-2">
          <button type="button" className="btn btn-primary" onClick={addSamples}>
            Isi data contoh
          </button>
        </div>
      </section>
    );
  }

  if (last) {
    const savedClient = clients.find((c) => c.id === last.client_id) ?? null;
    return (
      <section className="space-y-4">
        <div className="no-print panel flex flex-wrap items-center justify-between gap-3 p-4">
          <p className="text-sm">
            Tanda terima <strong className="tabular-nums">{last.nomor}</strong> tersimpan. Siap dicetak.
          </p>
          <button
            type="button"
            className="btn btn-line"
            onClick={() => {
              const n = nowJakarta();
              setClock({ date: n.date, time: n.time });
              setRows([{ key: `n${Date.now()}`, item: blankItem() }]);
              setCatatan("");
              setLast(null);
              setError("");
            }}
          >
            Tutup dan buat baru
          </button>
        </div>
        <div className="no-print space-y-3">
          <ModePicker mode={mode} onChange={setMode} />
          <PrintActions id={last.id} mode={mode} />
        </div>
        <SlipPreview
          mode={mode}
          slip={{
            nomor: last.nomor,
            tanggal: last.tanggal,
            jam: last.jam,
            up: last.up,
            catatan: last.catatan,
            items: last.items,
            client: savedClient,
            kurir: courierName(couriers, last.courier_id),
            kop,
          }}
        />
      </section>
    );
  }

  function patch(i: number, part: Partial<Item>) {
    setRows((list) => list.map((row, idx) => (idx === i ? { ...row, item: { ...row.item, ...part } } : row)));
  }

  const filled = rows.map((r) => r.item).filter((it) => it.Uraian.trim());

  return (
    <div className="grid items-start gap-6 xl:grid-cols-[minmax(0,22rem)_minmax(0,1fr)]">
      <section className="no-print space-y-4">
        <div>
          <h1 className="text-xl font-semibold">Buat tanda terima</h1>
          <p className="mt-1 text-sm text-muted">Nomor {previewNo} ditetapkan saat disimpan.</p>
        </div>
        <div className="panel space-y-3 p-4">
          <label className="block">
            <span className="lbl">Jenis</span>
            <select className="field" value={jenis} onChange={(e) => setJenis(e.target.value as Jenis)}>
              {JENIS_KEYS.map((k) => (
                <option key={k} value={k}>
                  {k} — {JENIS[k]}
                </option>
              ))}
            </select>
          </label>
          <div className="grid grid-cols-2 gap-3">
            <label className="block">
              <span className="lbl">Tanggal</span>
              <input className="field" type="date" value={clock.date} onChange={(e) => setClock({ ...clock, date: e.target.value })} />
            </label>
            <label className="block">
              <span className="lbl">Jam</span>
              <input className="field" type="time" value={clock.time} onChange={(e) => setClock({ ...clock, time: e.target.value })} />
            </label>
          </div>
          <label className="block">
            <span className="lbl">Pengantar</span>
            <select
              className="field"
              value={courier.id}
              onChange={(e) => setCourierId(Number(e.target.value))}
            >
              {active.map((c) => (
                <option key={c.id} value={c.id}>
                  {c.nama}
                </option>
              ))}
            </select>
          </label>
          <label className="block">
            <span className="lbl">Kepada</span>
            <select className="field" value={client.id} onChange={(e) => {
              const id = Number(e.target.value);
              setClientId(id);
              const next = sorted.find((c) => c.id === id);
              setUp(contactsOf(next)[0] ?? "");
            }}>
              {sorted.map((c) => (
                <option key={c.id} value={c.id}>
                  {c.nama}
                </option>
              ))}
            </select>
          </label>
          <p className="text-xs leading-relaxed text-muted">
            {[client.alamat1, client.alamat2, client.alamat3, client.alamat4].filter(Boolean).join(" · ") || "Alamat belum diisi."}
          </p>
          <label className="block">
            <span className="lbl">U.p.</span>
            <input className="field" list="up-list" value={up} onChange={(e) => setUp(e.target.value)} placeholder="Kosongkan jika tidak perlu" />
            <datalist id="up-list">
              {ups.map((name) => (
                <option key={name} value={name} />
              ))}
            </datalist>
          </label>
        </div>

        <div className="panel space-y-3 p-4">
          <div className="flex items-center justify-between gap-2">
            <h2 className="text-sm font-semibold">Daftar dokumen</h2>
            <button
              type="button"
              className="btn btn-line"
              onClick={() => setRows((list) => [...list, { key: `k${Date.now()}`, item: blankItem() }])}
            >
              <Plus className="size-4" aria-hidden />
              Baris
            </button>
          </div>
          {rows.map((row, i) => (
            <div key={row.key} className="grid gap-2 border-t border-line pt-3 first:border-t-0 first:pt-0">
              <label className="block">
                <span className="lbl">Uraian</span>
                <input className="field" value={row.item.Uraian} onChange={(e) => patch(i, { Uraian: e.target.value })} />
              </label>
              <div className="grid grid-cols-[minmax(0,1fr)_auto] gap-2">
                <label className="block">
                  <span className="lbl">Jumlah</span>
                  <input className="field" value={row.item.Jumlah} onChange={(e) => patch(i, { Jumlah: e.target.value })} />
                </label>
                <button
                  type="button"
                  className="btn btn-line mt-5"
                  aria-label="Hapus baris"
                  disabled={rows.length === 1}
                  onClick={() => setRows((list) => list.filter((_, idx) => idx !== i))}
                >
                  <Trash2 className="size-4" aria-hidden />
                </button>
              </div>
              <label className="block">
                <span className="lbl">Keterangan</span>
                <input className="field" value={row.item.Keterangan} onChange={(e) => patch(i, { Keterangan: e.target.value })} />
              </label>
            </div>
          ))}
          <label className="block">
            <span className="lbl">Catatan</span>
            <input className="field" value={catatan} onChange={(e) => setCatatan(e.target.value)} placeholder="Opsional" />
          </label>
          {error ? <p className="text-sm font-medium text-navy">{error}</p> : null}
          <button
            type="button"
            className="btn btn-primary w-full"
            onClick={() => {
              if (!filled.length) {
                setError("Isi minimal satu baris uraian.");
                return;
              }
              if (!clock.date || !clock.time) {
                setError("Tanggal dan jam wajib diisi.");
                return;
              }
              const rec = saveReceipt({
                jenis,
                tanggal: clock.date,
                jam: clock.time,
                client_id: client.id,
                up: up.trim(),
                courier_id: courier.id,
                catatan: catatan.trim(),
                items: filled,
              });
              setError("");
              setLast(rec);
              setMode("tt");
            }}
          >
            Simpan
          </button>
        </div>
      </section>

      <section className="min-w-0">
        <p className="no-print mb-3 text-xs text-muted">Pratinjau template. Tombol cetak muncul setelah disimpan.</p>
        <SlipPreview
          mode="tt"
          slip={{
            nomor: previewNo,
            tanggal: clock.date,
            jam: clock.time,
            up,
            catatan,
            items: filled,
            client,
            kurir: courier.nama,
            kop,
          }}
        />
      </section>
    </div>
  );
}

function Riwayat() {
  const receipts = useDesk((s) => s.receipts);
  const clients = useDesk((s) => s.clients);
  const couriers = useDesk((s) => s.couriers);
  const [q, setQ] = useState("");
  const [jenisOn, setJenisOn] = useState<Jenis[]>([]);
  const [statusOn, setStatusOn] = useState<Status[]>([]);
  const [sel, setSel] = useState<number | null>(null);

  const filtered = useMemo(() => {
    const needle = q.trim().toLowerCase();
    return [...receipts]
      .filter((r) => {
        if (jenisOn.length && !jenisOn.includes(r.jenis)) return false;
        if (statusOn.length && !statusOn.includes(r.status)) return false;
        if (!needle) return true;
        const hay = [
          r.nomor,
          r.tanggal,
          r.up,
          r.catatan,
          clientName(clients, r.client_id),
          courierName(couriers, r.courier_id),
          ...r.items.flatMap((it) => [it.Uraian, it.Jumlah, it.Keterangan]),
        ]
          .join(" ")
          .toLowerCase();
        return hay.includes(needle);
      })
      .sort((a, b) => b.tahun - a.tahun || b.seq - a.seq || b.id - a.id);
  }, [receipts, clients, couriers, q, jenisOn, statusOn]);

  const current = filtered.find((r) => r.id === sel) ?? filtered[0] ?? null;
  const dups = duplicateNomors(receipts);

  return (
    <section className="space-y-4">
      <h1 className="no-print text-xl font-semibold">Riwayat</h1>
      {dups.length ? (
        <p className="no-print panel flex items-start gap-2 p-3 text-sm">
          <AlertTriangle className="mt-0.5 size-4 shrink-0" aria-hidden />
          Nomor ganda pada tahun yang sama: {dups.join(", ")}.
        </p>
      ) : null}
      <div className="no-print space-y-3">
        <label className="block">
          <span className="lbl">Cari</span>
          <input className="field" value={q} onChange={(e) => setQ(e.target.value)} placeholder="Klien, nomor, atau isi dokumen" />
        </label>
        <div className="flex flex-wrap gap-2">
          {JENIS_KEYS.map((k) => (
            <button
              key={k}
              type="button"
              aria-pressed={jenisOn.includes(k)}
              className="btn btn-line"
              onClick={() => setJenisOn((list) => (list.includes(k) ? list.filter((x) => x !== k) : [...list, k]))}
            >
              {k}
            </button>
          ))}
          {STATUS.map((s) => (
            <button
              key={s}
              type="button"
              aria-pressed={statusOn.includes(s)}
              className="btn btn-line"
              onClick={() => setStatusOn((list) => (list.includes(s) ? list.filter((x) => x !== s) : [...list, s]))}
            >
              {s}
            </button>
          ))}
        </div>
      </div>
      {filtered.length === 0 ? (
        <p className="no-print text-sm text-muted">Belum ada tanda terima pada saringan ini.</p>
      ) : (
        <>
          <div className="no-print panel overflow-x-auto">
            <table className="w-full min-w-[40rem] text-left text-sm">
              <thead className="text-xs uppercase tracking-wide text-muted">
                <tr>
                  {["Nomor", "Jenis", "Tanggal", "Jam", "Klien", "Pengantar", "Status"].map((h) => (
                    <th key={h} className="px-3 py-2 font-semibold">{h}</th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {filtered.map((r) => (
                  <tr
                    key={r.id}
                    onClick={() => setSel(r.id)}
                    aria-selected={current?.id === r.id}
                    className={`cursor-pointer border-t border-line ${current?.id === r.id ? "bg-paper" : ""}`}
                  >
                    <td className="px-3 py-2 font-medium tabular-nums">{r.nomor}</td>
                    <td className="px-3 py-2">{r.jenis}</td>
                    <td className="px-3 py-2 tabular-nums">{r.tanggal}</td>
                    <td className="px-3 py-2 tabular-nums">{r.jam}</td>
                    <td className="px-3 py-2">{clientName(clients, r.client_id)}</td>
                    <td className="px-3 py-2">{courierName(couriers, r.courier_id) || "—"}</td>
                    <td className="px-3 py-2">{r.status}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
          {current ? <ReceiptDetail key={current.id} receipt={current} /> : null}
        </>
      )}
    </section>
  );
}

function ReceiptDetail({ receipt }: { receipt: Receipt }) {
  const clients = useDesk((s) => s.clients);
  const couriers = useDesk((s) => s.couriers);
  const kop = useDesk((s) => s.kop);
  const setStatus = useDesk((s) => s.setStatus);
  const deleteReceipt = useDesk((s) => s.deleteReceipt);
  const [status, setLocal] = useState<Status>(receipt.status);
  const [mode, setMode] = useState<PrintMode>("tt");
  const [arm, setArm] = useState(false);
  const missing = !clients.some((c) => c.id === receipt.client_id);

  return (
    <div className="space-y-4">
      <div className="no-print panel space-y-3 p-4">
        {missing ? (
          <p className="text-sm">Klien sudah tidak ada di master. Formulir tetap bisa dicetak dengan baris kepada kosong.</p>
        ) : null}
        <div className="flex flex-wrap items-end gap-2">
          <label className="block min-w-40">
            <span className="lbl">Status</span>
            <select className="field" value={status} onChange={(e) => setLocal(e.target.value as Status)}>
              {STATUS.map((s) => (
                <option key={s}>{s}</option>
              ))}
            </select>
          </label>
          <button type="button" className="btn btn-line" onClick={() => setStatus(receipt.id, status)}>
            Perbarui status
          </button>
          {arm ? (
            <button type="button" className="btn btn-line" onClick={() => deleteReceipt(receipt.id)}>
              Ya, hapus
            </button>
          ) : (
            <button type="button" className="btn btn-line" onClick={() => setArm(true)}>
              <Trash2 className="size-4" aria-hidden />
              Hapus
            </button>
          )}
        </div>
        <ModePicker mode={mode} onChange={setMode} />
        <PrintActions id={receipt.id} mode={mode} />
      </div>
      <SlipPreview
        mode={mode}
        slip={{
          nomor: receipt.nomor,
          tanggal: receipt.tanggal,
          jam: receipt.jam,
          up: receipt.up,
          catatan: receipt.catatan,
          items: receipt.items,
          client: clients.find((c) => c.id === receipt.client_id) ?? null,
          kurir: courierName(couriers, receipt.courier_id),
          kop,
        }}
      />
    </div>
  );
}

function Dasbor() {
  const receipts = useDesk((s) => s.receipts);
  const clients = useDesk((s) => s.clients);
  const couriers = useDesk((s) => s.couriers);
  const staleDays = useDesk((s) => s.staleDays);
  const setStaleDays = useDesk((s) => s.setStaleDays);
  const today = nowJakarta().date;
  const cutoff = addDays(today, -staleDays);
  const month = today.slice(0, 7);
  const overdue = receipts.filter((r) => r.status === "Diserahkan" && r.tanggal < cutoff);
  const byCourier = new Map<string, number>();
  for (const r of receipts) {
    const name = courierName(couriers, r.courier_id) || "—";
    byCourier.set(name, (byCourier.get(name) ?? 0) + 1);
  }
  const buckets = new Map<string, Record<Jenis, number>>();
  for (const r of receipts) {
    const key = r.tanggal.slice(0, 7);
    if (key.length !== 7) continue;
    const row = buckets.get(key) ?? { PR: 0, INV: 0, KW: 0, LP: 0 };
    row[r.jenis] += 1;
    buckets.set(key, row);
  }
  const series = [...buckets.entries()]
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([bulan, n]) => ({ bulan, ...n }));
  const maxC = Math.max(1, ...byCourier.values());

  if (!receipts.length) {
    return (
      <section>
        <h1 className="text-xl font-semibold">Dasbor</h1>
        <p className="mt-2 text-sm text-muted">Belum ada tanda terima.</p>
      </section>
    );
  }

  return (
    <section className="no-print space-y-5">
      <div className="flex flex-wrap items-end justify-between gap-3">
        <h1 className="text-xl font-semibold">Dasbor</h1>
        <label className="block w-40">
          <span className="lbl">Belum kembali setelah</span>
          <input
            className="field"
            type="number"
            min={1}
            max={90}
            value={staleDays}
            onChange={(e) => setStaleDays(Number(e.target.value) || 1)}
          />
        </label>
      </div>
      <div className="grid gap-3 sm:grid-cols-3">
        {[
          ["Total", receipts.length],
          ["Bulan ini", receipts.filter((r) => r.tanggal.startsWith(month)).length],
          ["Belum kembali", overdue.length],
        ].map(([label, n]) => (
          <div key={String(label)} className="panel p-4">
            <p className="text-xs uppercase tracking-wide text-muted">{label}</p>
            <p className="mt-1 text-3xl font-semibold tabular-nums">{n}</p>
          </div>
        ))}
      </div>
      <div className="panel p-4">
        <h2 className="text-sm font-semibold">Per bulan</h2>
        <div className="mt-3 h-56">
          <ResponsiveContainer width="100%" height="100%">
            <BarChart data={series}>
              <XAxis dataKey="bulan" tick={{ fontSize: 12 }} />
              <YAxis allowDecimals={false} width={28} tick={{ fontSize: 12 }} />
              <Tooltip />
              {CHART.map((c) => (
                <Bar key={c.key} dataKey={c.key} name={JENIS[c.key]} stackId="jenis" fill={c.fill} fillOpacity={c.opacity} />
              ))}
            </BarChart>
          </ResponsiveContainer>
        </div>
        <ul className="mt-3 flex flex-wrap gap-x-4 gap-y-1 text-xs text-muted">
          {CHART.map((c) => (
            <li key={c.key} className="flex items-center gap-1.5">
              <span className="inline-block size-2.5 rounded-sm" style={{ background: c.fill, opacity: c.opacity }} />
              {c.key} · {JENIS[c.key]}
            </li>
          ))}
        </ul>
      </div>
      <div className="panel p-4">
        <h2 className="mb-3 text-sm font-semibold">Per pengantar</h2>
        <ul className="space-y-2">
          {[...byCourier.entries()].sort((a, b) => b[1] - a[1]).map(([name, n]) => (
            <li key={name} className="grid grid-cols-[8rem_minmax(0,1fr)_2rem] items-center gap-3 text-sm">
              <span className="truncate">{name}</span>
              <span className="h-2 rounded-full bg-line">
                <span className="block h-2 rounded-full bg-navy" style={{ width: `${(n / maxC) * 100}%` }} />
              </span>
              <span className="tabular-nums text-right">{n}</span>
            </li>
          ))}
        </ul>
      </div>
      {overdue.length ? (
        <div className="panel p-4">
          <h2 className="text-sm font-semibold">Belum kembali</h2>
          <ul className="mt-2 divide-y divide-line text-sm">
            {overdue.map((r) => (
              <li key={r.id} className="flex flex-wrap justify-between gap-2 py-2">
                <span className="font-medium tabular-nums">{r.nomor}</span>
                <span className="text-muted">{r.tanggal}</span>
                <span>{clientName(clients, r.client_id)}</span>
                <span>{courierName(couriers, r.courier_id) || "—"}</span>
              </li>
            ))}
          </ul>
        </div>
      ) : null}
    </section>
  );
}

function KlienPage() {
  return (
    <div className="space-y-8">
      <ClientBook />
      <CourierBook />
    </div>
  );
}

function ClientBook() {
  const clients = useDesk((s) => s.clients);
  const saveClient = useDesk((s) => s.saveClient);
  const deleteClient = useDesk((s) => s.deleteClient);
  const sorted = useMemo(() => [...clients].sort((a, b) => a.nama.localeCompare(b.nama, "id")), [clients]);
  const [sel, setSel] = useState<number | "new">("new");
  const [draft, setDraft] = useState<Client>({ ...EMPTY_CLIENT });
  const [msg, setMsg] = useState("");

  useEffect(() => {
    if (sel === "new") setDraft({ ...EMPTY_CLIENT });
    else {
      const found = useDesk.getState().clients.find((c) => c.id === sel);
      setDraft(found ? { ...found } : { ...EMPTY_CLIENT });
    }
    setMsg("");
  }, [sel]);

  function set<K extends keyof Client>(key: K, value: Client[K]) {
    setDraft((d) => ({ ...d, [key]: value }));
  }

  const fields: { key: keyof Client; label: string; area?: boolean }[] = [
    { key: "nama", label: "Nama" },
    { key: "kode", label: "Kode" },
    { key: "alamat1", label: "Alamat 1" },
    { key: "alamat2", label: "Alamat 2" },
    { key: "alamat3", label: "Alamat 3" },
    { key: "alamat4", label: "Alamat 4" },
    { key: "kontak", label: "U.p. — satu nama per baris", area: true },
  ];

  return (
    <section className="no-print space-y-3">
      <h1 className="text-xl font-semibold">Klien</h1>
      <div className="grid gap-4 lg:grid-cols-[16rem_minmax(0,1fr)]">
        <div className="panel max-h-[28rem] overflow-auto p-2">
          <button type="button" className="btn btn-line mb-1 w-full" onClick={() => setSel("new")}>
            <Plus className="size-4" aria-hidden />
            Klien baru
          </button>
          {sorted.map((c) => (
            <button
              key={c.id}
              type="button"
              aria-current={sel === c.id ? "page" : undefined}
              className="btn btn-line mb-1 w-full justify-start"
              onClick={() => setSel(c.id)}
            >
              <span className="truncate">{c.nama}</span>
            </button>
          ))}
        </div>
        <form
          className="panel space-y-3 p-4"
          onSubmit={(e) => {
            e.preventDefault();
            if (!draft.nama.trim()) {
              setMsg("Nama klien wajib diisi.");
              return;
            }
            saveClient({ ...draft, nama: draft.nama.trim() });
            setDraft({ ...EMPTY_CLIENT });
            setMsg("Tersimpan.");
            setSel("new");
          }}
        >
          {fields.map((f) => (
            <label key={f.key} className="block">
              <span className="lbl">{f.label}</span>
              {f.area ? (
                <textarea className="field min-h-24" value={String(draft[f.key])} onChange={(e) => set(f.key, e.target.value)} />
              ) : (
                <input className="field" value={String(draft[f.key])} onChange={(e) => set(f.key, e.target.value)} />
              )}
            </label>
          ))}
          {msg ? <p className="text-sm text-navy">{msg}</p> : null}
          <div className="flex flex-wrap gap-2">
            <button type="submit" className="btn btn-primary">Simpan klien</button>
            {sel !== "new" ? (
              <button
                type="button"
                className="btn btn-line"
                onClick={() => {
                  const err = deleteClient(sel);
                  setMsg(err ?? "Klien dihapus.");
                  if (!err) setSel("new");
                }}
              >
                Hapus
              </button>
            ) : null}
          </div>
        </form>
      </div>
    </section>
  );
}

function CourierBook() {
  const couriers = useDesk((s) => s.couriers);
  const saveCourier = useDesk((s) => s.saveCourier);
  const deleteCourier = useDesk((s) => s.deleteCourier);
  const [nama, setNama] = useState("");
  const [msg, setMsg] = useState("");

  return (
    <section className="no-print space-y-3">
      <h2 className="text-xl font-semibold">Pengantar</h2>
      <ul className="panel divide-y divide-line">
        {couriers.map((c) => (
          <li key={c.id} className="flex flex-wrap items-center gap-2 p-3">
            <input
              className="field min-w-40 flex-1"
              value={c.nama}
              aria-label={`Nama ${c.nama}`}
              onChange={(e) => saveCourier({ ...c, nama: e.target.value })}
            />
            <button
              type="button"
              className="btn btn-line"
              aria-pressed={c.aktif}
              onClick={() => saveCourier({ ...c, aktif: !c.aktif })}
            >
              {c.aktif ? "Aktif" : "Nonaktif"}
            </button>
            <button
              type="button"
              className="btn btn-line"
              onClick={() => setMsg(deleteCourier(c.id) ?? "Pengantar dihapus.")}
            >
              Hapus
            </button>
          </li>
        ))}
      </ul>
      <form
        className="flex flex-wrap gap-2"
        onSubmit={(e) => {
          e.preventDefault();
          if (!nama.trim()) return;
          saveCourier({ id: 0, nama: nama.trim(), aktif: true });
          setNama("");
          setMsg("Pengantar ditambahkan.");
        }}
      >
        <input className="field max-w-xs" value={nama} onChange={(e) => setNama(e.target.value)} placeholder="Nama pengantar" aria-label="Nama pengantar baru" />
        <button type="submit" className="btn btn-primary">Tambah</button>
      </form>
      {msg ? <p className="text-sm text-navy">{msg}</p> : null}
    </section>
  );
}

function Atur() {
  const kop = useDesk((s) => s.kop);
  const setKop = useDesk((s) => s.setKop);
  const addSamples = useDesk((s) => s.addSamples);
  const replaceAll = useDesk((s) => s.replaceAll);
  const wipe = useDesk((s) => s.wipe);
  const [draft, setDraft] = useState(kop);
  const [msg, setMsg] = useState("");
  const [pending, setPending] = useState<{ data: DeskData; skipped: number } | null>(null);
  const [armWipe, setArmWipe] = useState(false);

  useEffect(() => setDraft(kop), [kop]);

  return (
    <section className="no-print max-w-2xl space-y-6">
      <div>
        <h1 className="text-xl font-semibold">Pengaturan</h1>
        <p className="mt-1 text-sm leading-relaxed text-muted">
          Yang bisa diubah hanya isi kop. Posisi, garis, kolom, dan judul TANDA TERIMA tetap seperti template.
        </p>
      </div>
      <form
        className="panel space-y-3 p-4"
        onSubmit={(e) => {
          e.preventDefault();
          setKop(draft);
          setMsg("Kop tersimpan.");
        }}
      >
        <h2 className="text-sm font-semibold">Kop surat</h2>
        <label className="block">
          <span className="lbl">Nama KJPP</span>
          <input className="field" value={draft.nama} onChange={(e) => setDraft({ ...draft, nama: e.target.value })} />
        </label>
        <label className="block">
          <span className="lbl">Blok kiri</span>
          <textarea className="field min-h-28" value={draft.kiri} onChange={(e) => setDraft({ ...draft, kiri: e.target.value })} />
        </label>
        <label className="block">
          <span className="lbl">Blok kanan</span>
          <textarea className="field min-h-36" value={draft.kanan} onChange={(e) => setDraft({ ...draft, kanan: e.target.value })} />
        </label>
        <div className="flex flex-wrap gap-2">
          <button type="submit" className="btn btn-primary">Simpan kop</button>
          <button type="button" className="btn btn-line" onClick={() => { setDraft({ ...KOP }); setKop({ ...KOP }); setMsg("Kop bawaan dikembalikan."); }}>
            Kembalikan bawaan
          </button>
        </div>
      </form>

      <div className="panel space-y-3 p-4">
        <h2 className="text-sm font-semibold">Cadangan</h2>
        <p className="text-sm leading-relaxed text-muted">
          JSON ini bisa dipulihkan di sini, dan bentuknya sama dengan cadangan aplikasi Streamlit.
        </p>
        <button
          type="button"
          className="btn btn-line"
          onClick={() => {
            const s = useDesk.getState();
            const blob = new Blob(
              [toBackup({ clients: s.clients, couriers: s.couriers, receipts: s.receipts, kop: s.kop, staleDays: s.staleDays })],
              { type: "application/json" },
            );
            const url = URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.href = url;
            a.download = `cadangan-tanda-terima-${nowJakarta().date.replace(/-/g, "")}.json`;
            a.click();
            URL.revokeObjectURL(url);
          }}
        >
          <Download className="size-4" aria-hidden />
          Unduh cadangan
        </button>
        <label className="block">
          <span className="lbl">Pulihkan dari cadangan</span>
          <input
            className="field"
            type="file"
            accept="application/json,.json"
            onChange={async (e) => {
              const file = e.target.files?.[0];
              e.target.value = "";
              if (!file) return;
              try {
                const parsed = parseBackup(JSON.parse(await file.text()));
                if (!parsed.ok) {
                  setPending(null);
                  setMsg(parsed.error);
                  return;
                }
                setPending({ data: parsed.data, skipped: parsed.skipped });
                setMsg(parsed.skipped ? `${parsed.skipped} baris rusak dilewati.` : "Berkas siap dipulihkan. Data saat ini akan tertimpa.");
              } catch {
                setPending(null);
                setMsg("JSON tidak bisa dibaca.");
              }
            }}
          />
        </label>
        {pending ? (
          <button
            type="button"
            className="btn btn-primary"
            onClick={() => {
              replaceAll(pending.data);
              setPending(null);
              setMsg("Data dipulihkan.");
            }}
          >
            Pulihkan sekarang
          </button>
        ) : null}
      </div>

      <div className="panel space-y-3 p-4">
        <h2 className="text-sm font-semibold">Data di peramban ini</h2>
        <button type="button" className="btn btn-line" onClick={() => { addSamples(); setMsg("Data contoh ditambahkan bila belum ada."); }}>
          Isi data contoh
        </button>
        {armWipe ? (
          <button type="button" className="btn btn-line" onClick={() => { wipe(); setArmWipe(false); setMsg("Semua data lokal dihapus. Kop kembali ke bawaan."); }}>
            Ya, hapus semua
          </button>
        ) : (
          <button type="button" className="btn btn-line" onClick={() => setArmWipe(true)}>
            Hapus semua data lokal
          </button>
        )}
      </div>
      {msg ? <p className="text-sm text-navy">{msg}</p> : null}
    </section>
  );
}
