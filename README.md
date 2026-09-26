# Server Sizer — VMware refresh build planner

Size a VMware refresh before you quote it. Feed in an RVTools export (or type
cluster numbers in manually), configure each cluster's target hosts, and get a
sized build: host specs, host counts, per-core licensing math, SE talking
points, and a customer-ready HTML report.

**Live:** https://scribnetai.github.io/server-sizer/

## How it works

1. **Load demand** — drop an RVTools `.xlsx` (File → Export all to Excel),
   enter cluster totals manually, or try the synthetic demo environment.
2. **Configure each cluster** — the mid-step: per-cluster target host specs
   (platform presets: Dell, Cisco UCS, HPE, Nutanix, Supermicro, or custom),
   growth headroom, CPU/RAM overcommit, N+1/N+2 redundancy, and HCI storage.
   Host-count previews update live as you tune.
3. **Build plan** — worked sizing math per cluster (demand → per-host usable
   capacity → binding constraint → spares), a bill-of-materials summary,
   vSphere per-core licensing with phantom-core waste, auto-written findings,
   and a downloadable standalone HTML briefing report.

## Sizing math

```
cpu hosts = ceil( demand vCPU ÷ (host cores × cpu overcommit × (1 − growth)) )
ghz hosts = ceil( demand vCPU × ghz/vCPU ÷ (host cores × base clock × (1 − growth)) )
mem hosts = ceil( demand vRAM ÷ (host RAM × ram overcommit × (1 − growth)) )
build     = max(cpu hosts, ghz hosts, mem hosts, storage hosts) + redundancy spares
```

Target CPUs are real Intel Xeon 6 SKUs (base clock used for sizing; turbo
ignored). The GHz dimension checks assumed sustained clock per vCPU
(tunable, default 0.5) against physical host clocks — no overcommit on the
host side, since overcommit is already expressed in the vCPU ratio.

Demand basis is **allocated** resources or **actual utilization** (from the
export's host utilization figures). Licensing follows Broadcom's per-core
model: every socket bills `max(cores per socket, 16)`.

## Privacy

100% client-side. The spreadsheet parser (vendored SheetJS) runs in the page,
nothing is uploaded, nothing is stored (no localStorage/IndexedDB/cookies),
and the page works fully offline after load.

## Honest limitations

- Capacity sizing, not performance sizing — no IOPS/latency/NUMA modeling.
- RVTools utilization is point-in-time, not an average.
- Overcommit ratios are planner assumptions, not measurements.
- Licensing covers vSphere compute cores only (no vSAN/NSX/Aria/Windows).
- No pricing — output is host counts, specs, and license cores for your
  VAR quoting tools.

## Files

- `index.html` — landing page + 3-step wizard
- `css/styles.css` — theme (cache-busted via `?v=N`)
- `js/app.js` — parsing, sizing math, wizard, report generator
- `lib/xlsx.full.min.js` — vendored SheetJS (offline parsing)

Not affiliated with Broadcom/VMware.
