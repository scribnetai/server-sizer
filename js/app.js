'use strict';
/* ============================================================
   Server Sizer — VMware refresh build planner
   100% client-side. No uploads, no storage, no network calls
   carrying customer data. Everything lives in page memory.
   ============================================================ */

/* ================= Helpers ================= */
const $ = (id) => document.getElementById(id);
const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const fmtInt = (n) => Math.round(n || 0).toLocaleString('en-US');
const fmt1 = (n) => (n == null || !isFinite(n)) ? '—' : (Math.round(n * 10) / 10).toLocaleString('en-US');
const fmtTB = (tb) => (tb == null || !isFinite(tb)) ? '—' : (tb >= 100 ? fmtInt(tb) : fmt1(tb)) + ' TB';
function parseNum(v) {
  if (v == null || v === '') return 0;
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}
function parsePct(v) {
  if (v == null || v === '') return null;
  const m = String(v).trim().match(/([\d.]+)\s*%/);
  if (m) return parseFloat(m[1]) / 100;
  const n = parseFloat(String(v));
  if (!isFinite(n)) return null;
  return n > 1 ? n / 100 : n;
}
function pick(row, aliases) {
  for (const a of aliases) {
    if (row[a] !== undefined && row[a] !== '' && row[a] != null) return row[a];
  }
  return '';
}
function lcg(seed) {
  let s = seed >>> 0;
  return () => { s = (s * 1664525 + 1013904223) >>> 0; return s / 4294967296; };
}

/* ================= RVTools column aliases (modern + legacy) ================= */
const COL = {
  vm: {
    name: ['VM'], power: ['Powerstate', 'Power state'], template: ['Template'],
    cpus: ['CPUs', 'Num CPUs'], memMB: ['Memory', 'Memory MB', 'Memory MiB'],
    usedMB: ['In Use MB', 'In use MiB', 'Used MB', 'Used MiB'],
    dc: ['Datacenter', 'DC'], cluster: ['Cluster'], host: ['Host', 'ESX host'],
  },
  host: {
    name: ['Host'], dc: ['Datacenter'], cluster: ['Cluster'],
    sockets: ['# CPU', '#CPU', 'Sockets', 'Num CPU'],
    coresPerCpu: ['Cores per CPU', 'Cores/CPU', 'Cores per cpu'],
    cores: ['# Cores', '#Cores', 'Total Cores'],
    cpuPct: ['CPU usage %', '% CPU'], memMB: ['# Memory', 'Memory MB', 'Memory size', 'Total Memory'],
    memPct: ['Memory usage %', '% Memory'],
  },
};
const TAB_NAMES = {
  vInfo: ['vInfo'], vHost: ['vHost'], vCluster: ['vCluster'], vDatastore: ['vDatastore'],
};
function findSheet(wb, candidates) {
  const lower = wb.SheetNames.map((s) => s.toLowerCase());
  for (const c of candidates) {
    const i = lower.indexOf(c.toLowerCase());
    if (i >= 0) return wb.SheetNames[i];
  }
  return null;
}
function sheetRows(ws) {
  if (!ws) return [];
  const asArray = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: true });
  if (!asArray.length) return [];
  let headerIdx = 0;
  for (let i = 0; i < Math.min(asArray.length, 5); i++) {
    const joined = asArray[i].join(' ').toLowerCase();
    if (/(^|\s)(vm|host|name|cluster|datastore|snapshot)(s|\s|$)/.test(' ' + joined + ' ')) { headerIdx = i; break; }
  }
  return XLSX.utils.sheet_to_json(ws, { defval: '', raw: true, range: headerIdx });
}
function normPower(v) {
  const s = String(v || '').toLowerCase().replace(/[\s_-]/g, '');
  if (s === 'poweredon') return 'on';
  if (s === 'poweredoff') return 'off';
  if (s === 'suspended') return 'suspended';
  return 'unknown';
}
function normVMs(rows) {
  return rows.map((r) => ({
    name: String(pick(r, COL.vm.name) || 'unknown'),
    power: normPower(pick(r, COL.vm.power)),
    template: /true|yes/i.test(String(pick(r, COL.vm.template))),
    cpus: parseNum(pick(r, COL.vm.cpus)),
    memMB: parseNum(pick(r, COL.vm.memMB)),
    usedMB: parseNum(pick(r, COL.vm.usedMB)),
    dc: String(pick(r, COL.vm.dc) || '—'),
    cluster: String(pick(r, COL.vm.cluster) || ''),
    host: String(pick(r, COL.vm.host) || ''),
  }));
}
function normHosts(rows) {
  return rows.map((r) => {
    const sockets = parseNum(pick(r, COL.host.sockets));
    let coresPerCpu = parseNum(pick(r, COL.host.coresPerCpu));
    let cores = parseNum(pick(r, COL.host.cores));
    if (!coresPerCpu && sockets && cores) coresPerCpu = cores / sockets;
    if (!cores && sockets && coresPerCpu) cores = sockets * coresPerCpu;
    return {
      name: String(pick(r, COL.host.name) || 'unknown'),
      dc: String(pick(r, COL.host.dc) || '—'),
      cluster: String(pick(r, COL.host.cluster) || ''),
      sockets, coresPerCpu, cores,
      cpuPct: parsePct(pick(r, COL.host.cpuPct)),
      memMB: parseNum(pick(r, COL.host.memMB)),
      memPct: parsePct(pick(r, COL.host.memPct)),
    };
  });
}

/* ================= Demand model ================= */
function clKey(dc, cluster) { return dc + ' / ' + (cluster || 'Standalone'); }

function buildClusters(parsed) {
  const byCl = {};
  parsed.vms.filter((v) => !v.template).forEach((v) => {
    const k = clKey(v.dc, v.cluster);
    if (!byCl[k]) byCl[k] = { name: k, vms: 0, poweredOn: 0, allocVCpu: 0, allocMemGB: 0, usedStorageTB: 0 };
    const c = byCl[k];
    c.vms++;
    if (v.power === 'on') { c.poweredOn++; c.allocVCpu += v.cpus; c.allocMemGB += v.memMB / 1024; }
    c.usedStorageTB += (v.usedMB || 0) / 1048576;
  });
  // Current hosts + utilization + current licensing, per cluster
  const hostAgg = {};
  parsed.hosts.forEach((h) => {
    const k = clKey(h.dc, h.cluster);
    if (!hostAgg[k]) hostAgg[k] = { n: 0, cpuW: 0, cpuBase: 0, memW: 0, memBase: 0, lic: 0, sockets: 0, cpsSum: 0 };
    const e = hostAgg[k];
    e.n++;
    const cores = h.cores || (h.sockets * h.coresPerCpu) || 0;
    if (h.cpuPct != null && cores > 0) { e.cpuW += h.cpuPct * cores; e.cpuBase += cores; }
    if (h.memPct != null && h.memMB > 0) { e.memW += h.memPct * h.memMB; e.memBase += h.memMB; }
    if (h.sockets > 0) {
      e.sockets += h.sockets;
      e.cpsSum += h.coresPerCpu || 0;
      e.lic += h.sockets * Math.max(h.coresPerCpu || 0, 16);
    }
  });
  let skipped = 0;
  const clusters = Object.keys(byCl).sort().map((k, i) => {
    const c = byCl[k];
    if (c.vms === 0) { skipped++; return null; }
    const ha = hostAgg[k] || { n: 0 };
    return {
      id: 'c' + i, name: k,
      vms: c.vms, poweredOn: c.poweredOn,
      allocVCpu: Math.round(c.allocVCpu), allocMemGB: Math.round(c.allocMemGB),
      usedStorageTB: Math.round(c.usedStorageTB * 10) / 10,
      avgCpuUtil: ha.cpuBase > 0 ? ha.cpuW / ha.cpuBase : null,
      avgMemUtil: ha.memBase > 0 ? ha.memW / ha.memBase : null,
      curHosts: ha.n || null,
      curHostCps: ha.n > 0 && ha.sockets > 0 ? Math.round(ha.cpsSum / ha.n * 10) / 10 : null,
      curLicenseCores: ha.lic > 0 ? ha.lic : null,
      source: 'rvtools',
    };
  }).filter(Boolean);
  return { clusters, skipped };
}

/* ================= Demo data (synthetic, in-memory only) ================= */
function genDemoClusters() {
  const rnd = lcg(20260926);
  const mk = (name, vms, vcpu, memGB, stoTB, hosts, hostCps, cpuU, memU) => ({
    id: 'd' + Math.floor(rnd() * 1e6), name, vms,
    poweredOn: Math.round(vms * (0.85 + rnd() * 0.1)),
    allocVCpu: vcpu, allocMemGB: memGB, usedStorageTB: stoTB,
    avgCpuUtil: cpuU, avgMemUtil: memU,
    curHosts: hosts, curHostCps: hostCps,
    curLicenseCores: hosts * 2 * Math.max(hostCps, 16),
    source: 'demo',
  });
  return [
    mk('DC-East / Prod-General', 184, 612, 2380, 46, 8, 10, 0.42, 0.58),
    mk('DC-East / Prod-DB', 58, 472, 3840, 31, 6, 12, 0.35, 0.71),
    mk('DC-West / ROBO-Edge', 26, 92, 410, 8, 3, 8, 0.28, 0.49),
  ];
}

/* ================= Sizing math ================= */
const CPUS = {
  'Xeon 6980P': { cores: 128, ghz: 2.0, family: 'Xeon 6 6900P' },
  'Xeon 6972P': { cores: 96, ghz: 2.4, family: 'Xeon 6 6900P' },
  'Xeon 6960P': { cores: 72, ghz: 2.7, family: 'Xeon 6 6900P' },
  'Xeon 6787P': { cores: 86, ghz: 2.0, family: 'Xeon 6 6700P' },
  'Xeon 6767P': { cores: 64, ghz: 2.4, family: 'Xeon 6 6700P' },
  'Xeon 6747P': { cores: 48, ghz: 2.7, family: 'Xeon 6 6700P' },
  'Xeon 6745P': { cores: 32, ghz: 3.1, family: 'Xeon 6 6700P' },
  'Xeon 6737P': { cores: 32, ghz: 2.9, family: 'Xeon 6 6700P' },
  'Xeon 6730P': { cores: 32, ghz: 2.5, family: 'Xeon 6 6700P' },
  'Xeon 6527P': { cores: 24, ghz: 3.0, family: 'Xeon 6 6500P' },
  'Xeon 6520P': { cores: 24, ghz: 2.4, family: 'Xeon 6 6500P' },
  'Xeon 6517P': { cores: 16, ghz: 3.2, family: 'Xeon 6 6500P' },
  'Xeon 6505P': { cores: 12, ghz: 2.2, family: 'Xeon 6 6500P' },
  'Xeon 6507P': { cores: 8, ghz: 3.5, family: 'Xeon 6 6500P' },
  'Xeon 6377P': { cores: 12, ghz: 3.1, family: 'Xeon 6300 entry' },
  'Xeon 6780E': { cores: 144, ghz: 2.2, family: 'Xeon 6 6700E' },
  'Xeon 6710E': { cores: 64, ghz: 2.4, family: 'Xeon 6 6700E' },
  'Custom': null,
};
const PLATFORMS = {
  'Dell PowerEdge R760': { sockets: 2, cpu: 'Xeon 6745P', ramGB: 1024 },
  'Cisco UCS C240 M7': { sockets: 2, cpu: 'Xeon 6745P', ramGB: 1024 },
  'HPE ProLiant DL380 Gen11': { sockets: 2, cpu: 'Xeon 6527P', ramGB: 768 },
  'Nutanix NX-8155N': { sockets: 2, cpu: 'Xeon 6745P', ramGB: 1024 },
  'Supermicro Hyper (value)': { sockets: 2, cpu: 'Xeon 6730P', ramGB: 512 },
  'Custom': null,
};
const STORAGE_REPL = { 'mirror': { label: 'FTT=1 mirror (2×)', factor: 2.0 }, 'raid5': { label: 'FTT=1 RAID-5/6 (1.33×)', factor: 1.33 } };

function defaultCfg() {
  return {
    basis: 'allocated', growth: 0.20, redundancy: 'n1',
    platform: 'Dell PowerEdge R760', sockets: 2, cps: 32, ramGB: 1024,
    cpu: 'Xeon 6745P', ghz: 3.1, ghzPerVcpu: 0.5,
    cpuOC: 4, memOC: 1.25, hci: false, storageTB: 15, storageRepl: 'mirror',
  };
}

function sizeCluster(cluster, cfg) {
  const hasUtil = cluster.avgCpuUtil != null && cluster.avgMemUtil != null;
  const useActual = cfg.basis === 'actual' && hasUtil;
  const demandVCpu = useActual ? cluster.allocVCpu * cluster.avgCpuUtil : cluster.allocVCpu;
  const demandMemGB = useActual ? cluster.allocMemGB * cluster.avgMemUtil : cluster.allocMemGB;
  const demandStoTB = cluster.usedStorageTB;
  const eff = 1 - cfg.growth;

  const perHostVCpu = cfg.sockets * cfg.cps * cfg.cpuOC;
  const perHostMemGB = cfg.ramGB * cfg.memOC;
  const cpuHosts = Math.max(1, Math.ceil(demandVCpu / (perHostVCpu * eff)));
  const memHosts = Math.max(1, Math.ceil(demandMemGB / (perHostMemGB * eff)));

  // GHz dimension: sustained clock demand vs physical clocks per host.
  // No overcommit on the host side — overcommit is already expressed in the
  // vCPU ratio; this checks whether the assumed GHz/vCPU fits the silicon.
  const demandGHz = demandVCpu * (cfg.ghzPerVcpu || 0.5);
  const perHostGHz = cfg.sockets * cfg.cps * (cfg.ghz || 2.5);
  const ghzHosts = Math.max(1, Math.ceil(demandGHz / (perHostGHz * eff)));

  let stoHosts = 0, perHostUsableTB = 0;
  if (cfg.hci) {
    const repl = STORAGE_REPL[cfg.storageRepl] ? STORAGE_REPL[cfg.storageRepl].factor : 2.0;
    perHostUsableTB = cfg.storageTB / repl;
    stoHosts = demandStoTB > 0 ? Math.max(1, Math.ceil(demandStoTB / (perHostUsableTB * eff))) : 1;
  }

  const rawHosts = Math.max(cpuHosts, ghzHosts, memHosts, stoHosts, 1);
  const binding = rawHosts === stoHosts && cfg.hci ? 'storage'
    : rawHosts === ghzHosts && rawHosts !== cpuHosts && rawHosts !== memHosts ? 'ghz'
    : rawHosts === memHosts && rawHosts !== cpuHosts ? 'memory'
    : rawHosts === cpuHosts && rawHosts !== memHosts ? 'cpu' : 'balanced';
  const spares = cfg.redundancy === 'n1' ? 1 : cfg.redundancy === 'n2' ? 2 : 0;
  const finalHosts = rawHosts + spares;

  const licPerHost = cfg.sockets * Math.max(cfg.cps, 16);
  const phantomPerHost = licPerHost - cfg.sockets * cfg.cps;
  const totalLic = finalHosts * licPerHost;
  const totalPhantom = finalHosts * phantomPerHost;
  const totalPhysCores = finalHosts * cfg.sockets * cfg.cps;

  const capCpuUsed = demandVCpu / (finalHosts * perHostVCpu);
  const capMemUsed = demandMemGB / (finalHosts * perHostMemGB);
  const capGhzUsed = demandGHz / (finalHosts * perHostGHz);

  return {
    useActual, demandVCpu, demandMemGB, demandStoTB, eff,
    perHostVCpu, perHostMemGB, perHostUsableTB,
    demandGHz, perHostGHz, ghzHosts,
    cpuHosts, memHosts, stoHosts, rawHosts, binding, spares, finalHosts,
    licPerHost, phantomPerHost, totalLic, totalPhantom, totalPhysCores,
    capCpuUsed, capMemUsed, capGhzUsed,
    hostLabel: cfg.sockets + '×' + cfg.cps + 'c ' + (cfg.cpu || 'Custom') + ' @ ' + (cfg.ghz || 2.5).toFixed(1) + 'GHz · ' + fmtInt(cfg.ramGB) + ' GB',
  };
}

/* ================= App state ================= */
const APP = { clusters: [], cfgs: {}, source: null, fileName: null, results: null };

function setStatus(msg) { const s = $('parseStatus'); s.hidden = false; s.innerHTML = msg; }
function showError(msg) { const e = $('fileError'); e.hidden = false; e.innerHTML = msg; }
function clearMsgs() { $('fileError').hidden = true; $('parseStatus').hidden = true; }

function setStep(n) {
  $('stepData').hidden = n !== 1;
  $('stepConfig').hidden = n !== 2;
  $('stepResults').hidden = n !== 3;
  document.querySelectorAll('#stepper .step').forEach((el) => {
    const s = parseInt(el.dataset.step);
    el.classList.toggle('active', s === n);
    el.classList.toggle('done', s < n);
  });
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

function startWizard() {
  $('landing').hidden = true;
  $('wizard').hidden = false;
  setStep(1);
  window.scrollTo({ top: 0 });
}

/* ================= Data loading ================= */
function parseWorkbook(wb, fileName) {
  const vms = normVMs(sheetRows(wb.Sheets[findSheet(wb, TAB_NAMES.vInfo)] || {}));
  const hws = findSheet(wb, TAB_NAMES.vHost);
  const hosts = hws ? normHosts(sheetRows(wb.Sheets[hws])) : [];
  if (!vms.length) {
    showError('<strong>No VM rows found.</strong> Make sure this is an RVTools export with a <code>vInfo</code> tab (File → Export all to Excel).');
    return;
  }
  const { clusters, skipped } = buildClusters({ vms, hosts });
  if (!clusters.length) { showError('<strong>No clusters with VMs found</strong> in this export.'); return; }
  APP.clusters = clusters;
  APP.source = 'rvtools';
  APP.fileName = fileName;
  setStatus('Parsed <strong>' + fmtInt(vms.length) + '</strong> VMs and <strong>' + fmtInt(hosts.length) + '</strong> hosts → <strong>' + clusters.length + '</strong> clusters.' + (skipped ? ' (' + skipped + ' empty skipped)' : ''));
  renderInventory();
  queueAutosave();
}

async function handleFile(file) {
  clearMsgs();
  if (!file) return;
  if (!/\.(xlsx|xls)$/i.test(file.name)) { showError('<strong>Not a spreadsheet.</strong> Drop the <code>.xlsx</code> from RVTools (File → Export all to Excel).'); return; }
  setStatus('Reading <strong>' + esc(file.name) + '</strong>…');
  try {
    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: 'array' });
    parseWorkbook(wb, file.name);
  } catch (err) {
    showError('<strong>Could not parse that file.</strong> ' + esc(err.message || 'Unknown error.'));
  }
}

function loadDemo() {
  clearMsgs();
  APP.clusters = genDemoClusters();
  APP.source = 'demo';
  APP.fileName = 'demo-environment';
  setStatus('Loaded <strong>demo environment</strong>: 3 synthetic clusters with different profiles (general, memory-heavy DB, ROBO edge). Generated in your browser — nothing uploaded.');
  renderInventory();
  queueAutosave();
}

/* ---- Manual entry ---- */
function manualRowHTML() {
  return '<tr>' +
    '<td><input data-f="name" placeholder="e.g. Prod-General" value=""></td>' +
    '<td><input class="num" data-f="vms" type="number" min="0" value=""></td>' +
    '<td><input class="num" data-f="vcpu" type="number" min="0" value=""></td>' +
    '<td><input class="num" data-f="ramGB" type="number" min="0" value=""></td>' +
    '<td><input class="num" data-f="stoTB" type="number" min="0" step="0.1" value=""></td>' +
    '<td><input class="num" data-f="curHosts" type="number" min="0" value=""></td>' +
    '<td><input class="num" data-f="curCps" type="number" min="0" value=""></td>' +
    '<td><button class="btn ghost" data-del style="padding:6px 10px">✕</button></td></tr>';
}
function wireManualEditor() {
  const body = $('manualBody');
  if (!body.children.length) { body.innerHTML = manualRowHTML() + manualRowHTML(); }
  body.querySelectorAll('[data-del]').forEach((b) => {
    b.onclick = () => { if (body.children.length > 1) b.closest('tr').remove(); };
  });
}
function applyManual() {
  clearMsgs();
  const rows = [...$('manualBody').querySelectorAll('tr')];
  const clusters = [];
  rows.forEach((tr, i) => {
    const g = (f) => tr.querySelector('[data-f="' + f + '"]').value.trim();
    const name = g('name') || ('Cluster ' + (i + 1));
    const vms = parseNum(g('vms')), vcpu = parseNum(g('vcpu')), ramGB = parseNum(g('ramGB')), stoTB = parseNum(g('stoTB'));
    if (!vms && !vcpu && !ramGB) return; // skip empty rows
    const curHosts = parseNum(g('curHosts')) || null;
    const curCps = parseNum(g('curCps')) || null;
    clusters.push({
      id: 'm' + i, name, vms, poweredOn: vms,
      allocVCpu: Math.round(vcpu), allocMemGB: Math.round(ramGB),
      usedStorageTB: Math.round(stoTB * 10) / 10,
      avgCpuUtil: null, avgMemUtil: null,
      curHosts, curHostCps: curCps,
      curLicenseCores: (curHosts && curCps) ? curHosts * 2 * Math.max(curCps, 16) : null, // assumes 2 sockets/host; curCps is cores per socket
      source: 'manual',
    });
  });
  if (!clusters.length) { showError('<strong>No usable rows.</strong> Fill in at least cluster name, VMs, vCPUs, and RAM for one cluster.'); return; }
  APP.clusters = clusters;
  APP.source = 'manual';
  APP.fileName = 'manual-entry';
  setStatus('Using <strong>' + clusters.length + '</strong> manually entered cluster' + (clusters.length > 1 ? 's' : '') + '. "Actual utilization" sizing basis is unavailable without export data — allocated totals are used.');
  renderInventory();
  queueAutosave();
}

function renderInventory() {
  const w = $('inventoryWrap');
  w.hidden = false;
  $('inventoryMeta').textContent = APP.source === 'rvtools' ? 'from ' + APP.fileName : APP.source === 'demo' ? 'synthetic demo data' : 'manual entry';
  $('inventoryBody').innerHTML = APP.clusters.map((c) => {
    const util = (c.avgCpuUtil != null)
      ? Math.round(c.avgCpuUtil * 100) + '% CPU · ' + Math.round(c.avgMemUtil * 100) + '% MEM'
      : '<span class="muted">n/a</span>';
    return '<tr><td><strong>' + esc(c.name) + '</strong></td>' +
      '<td class="num">' + fmtInt(c.vms) + '</td>' +
      '<td class="num">' + fmtInt(c.allocVCpu) + '</td>' +
      '<td class="num">' + fmtInt(c.allocMemGB) + '</td>' +
      '<td class="num">' + fmtTB(c.usedStorageTB) + '</td>' +
      '<td class="num">' + (c.curHosts != null ? c.curHosts : '—') + '</td>' +
      '<td>' + util + '</td></tr>';
  }).join('');
  w.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

/* ================= Step 2: per-cluster configuration ================= */
function getCfg(id) {
  if (!APP.cfgs[id]) APP.cfgs[id] = defaultCfg();
  const cfg = APP.cfgs[id];
  // Migrate projects saved before the GHz/CPU-model update.
  if (cfg.cpu == null) {
    const p = PLATFORMS[cfg.platform];
    if (p && p.cpu && CPUS[p.cpu] && cfg.cps === CPUS[p.cpu].cores) {
      cfg.cpu = p.cpu; cfg.ghz = CPUS[p.cpu].ghz;
    } else {
      cfg.cpu = 'Custom';
      if (cfg.ghz == null) cfg.ghz = 2.5;
    }
  }
  if (cfg.ghzPerVcpu == null) cfg.ghzPerVcpu = 0.5;
  return cfg;
}

function segHTML(seg, opts, cur) {
  return '<div class="seg" data-seg="' + seg + '">' + opts.map((o) =>
    '<button type="button" data-val="' + o[0] + '"' + (o[0] === cur ? ' class="active"' : '') + (o[2] ? ' disabled title="' + esc(o[2]) + '"' : '') + '>' + o[1] + '</button>'
  ).join('') + '</div>';
}

function renderConfig() {
  const wrap = $('clusterCards');
  wrap.innerHTML = APP.clusters.map((c, i) => {
    const cfg = getCfg(c.id);
    const hasUtil = c.avgCpuUtil != null;
    const platOpts = Object.keys(PLATFORMS).map((p) => '<option' + (p === cfg.platform ? ' selected' : '') + '>' + esc(p) + '</option>').join('');
    const cpuOpts = Object.keys(CPUS).map((k) => {
      const d = CPUS[k];
      const lbl = d ? k + ' — ' + d.cores + 'c @ ' + d.ghz.toFixed(1) + ' GHz' : k;
      return '<option value="' + esc(k) + '"' + (k === cfg.cpu ? ' selected' : '') + '>' + esc(lbl) + '</option>';
    }).join('');
    const cpsOpts = [8, 10, 12, 16, 20, 24, 28, 32, 36, 40, 48, 56, 64, 72, 86, 96, 120, 128, 144].map((v) => '<option value="' + v + '"' + (v === cfg.cps ? ' selected' : '') + '>' + v + '</option>').join('');
    const ramOpts = [128, 256, 384, 512, 768, 1024, 1536, 2048].map((v) => '<option value="' + v + '"' + (v === cfg.ramGB ? ' selected' : '') + '>' + fmtInt(v) + ' GB</option>').join('');
    return '<div class="ccard' + (i === 0 ? ' open' : '') + '" data-id="' + c.id + '">' +
      '<div class="ccard-head">' +
        '<div><div class="ccard-title">' + esc(c.name) + '</div>' +
        '<div class="ccard-sub">' + fmtInt(c.vms) + ' VMs · ' + fmtInt(c.allocVCpu) + ' vCPU · ' + fmtInt(c.allocMemGB) + ' GB RAM · ' + fmtTB(c.usedStorageTB) + '</div></div>' +
        '<div class="ccard-preview"><div class="hosts"><span data-pv="hosts">—</span> <small>hosts</small></div><div class="binding" data-pv="binding"></div></div>' +
        '<div class="ccard-toggle">▾</div>' +
      '</div>' +
      '<div class="ccard-body"><div class="cfg-grid">' +
        '<div class="cfg-field"><label>Demand basis</label>' +
          segHTML('basis', [['allocated', 'Allocated'], ['actual', 'Actual util', hasUtil ? '' : 'No utilization data for this cluster']], cfg.basis) +
          '<div class="cfg-note">' + (hasUtil ? 'Actual = allocated × avg utilization from the export (' + Math.round(c.avgCpuUtil * 100) + '% CPU · ' + Math.round(c.avgMemUtil * 100) + '% MEM).' : 'Utilization data unavailable — allocated totals used.') + '</div></div>' +
        '<div class="cfg-field"><label>Growth headroom: <strong data-lb="growth">' + Math.round(cfg.growth * 100) + '%</strong></label>' +
          '<input type="range" data-cfg="growth" min="0" max="50" step="5" value="' + Math.round(cfg.growth * 100) + '">' +
          '<div class="cfg-note">Usable capacity per host is derated by this much.</div></div>' +
        '<div class="cfg-field"><label>Redundancy</label>' +
          segHTML('redundancy', [['none', 'None'], ['n1', 'N+1'], ['n2', 'N+2']], cfg.redundancy) +
          '<div class="cfg-note">Spare hosts for failure + maintenance.</div></div>' +
        '<div class="cfg-field"><label>Platform preset</label>' +
          '<select data-cfg="platform">' + platOpts + '</select>' +
          '<div class="cfg-note">Presets fill sockets / CPU / RAM — tune freely after.</div></div>' +
        '<div class="cfg-field"><label>CPU model</label>' +
          '<select data-cfg="cpu">' + cpuOpts + '</select>' +
          '<div class="cfg-note">Current Intel Xeon 6 line — base clock is used for sizing (turbo ignored).</div></div>' +
        '<div class="cfg-field"><label>Sockets / cores / base clock (GHz)</label><div class="cfg-row">' +
          '<select data-cfg="sockets">' + [1, 2, 4].map((v) => '<option value="' + v + '"' + (v === cfg.sockets ? ' selected' : '') + '>' + v + '</option>').join('') + '</select>' +
          '<select data-cfg="cps">' + cpsOpts + '</select>' +
          '<input type="number" data-cfg="ghz" min="1" max="5" step="0.1" value="' + (cfg.ghz || 2.5).toFixed(1) + '" title="Base clock in GHz"></div>' +
          '<div class="cfg-note">Editing cores or clock flips the CPU to Custom. vSphere bills max(cores/socket, 16).</div></div>' +
        '<div class="cfg-field"><label>RAM per host</label>' +
          '<select data-cfg="ramGB">' + ramOpts + '</select></div>' +
        '<div class="cfg-field"><label>CPU overcommit: <strong data-lb="cpuOC">' + cfg.cpuOC.toFixed(1) + ':1</strong></label>' +
          '<input type="range" data-cfg="cpuOC" min="1" max="10" step="0.5" value="' + cfg.cpuOC + '">' +
          '<div class="cfg-note">vCPUs per physical core. Your judgment call.</div></div>' +
        '<div class="cfg-field"><label>Avg GHz per vCPU: <strong data-lb="ghzPerVcpu">' + (cfg.ghzPerVcpu || 0.5).toFixed(1) + '</strong></label>' +
          '<input type="range" data-cfg="ghzPerVcpu" min="0.1" max="3" step="0.1" value="' + (cfg.ghzPerVcpu || 0.5) + '">' +
          '<div class="cfg-note">Sustained clock assumed per vCPU. At 4:1 overcommit on 3.1 GHz cores each vCPU gets ~0.8 GHz — raise for clock-hungry workloads.</div></div>' +
        '<div class="cfg-field"><label>RAM overcommit: <strong data-lb="memOC">' + cfg.memOC.toFixed(2) + ':1</strong></label>' +
          '<input type="range" data-cfg="memOC" min="1" max="2" step="0.05" value="' + cfg.memOC + '">' +
          '<div class="cfg-note">Keep ≤1.5:1 unless you know the workload.</div></div>' +
        '<div class="cfg-field"><label>Architecture</label>' +
          segHTML('arch', [['compute', 'Compute-only'], ['hci', 'HCI']], cfg.hci ? 'hci' : 'compute') +
          '<div class="cfg-note">HCI adds a storage-driven host count.</div></div>' +
        '<div class="cfg-field" data-hci-only' + (cfg.hci ? '' : ' hidden') + '><label>Raw storage per host (TB)</label>' +
          '<input type="number" data-cfg="storageTB" min="1" step="1" value="' + cfg.storageTB + '"></div>' +
        '<div class="cfg-field" data-hci-only' + (cfg.hci ? '' : ' hidden') + '><label>Storage resilience</label>' +
          segHTML('storageRepl', [['mirror', 'FTT=1 mirror'], ['raid5', 'RAID-5/6']], cfg.storageRepl) +
          '<div class="cfg-note">Usable per host = raw ÷ replication.</div></div>' +
      '</div>' +
      '<div class="spec-line" data-pv="spec"></div>' +
      '</div></div>';
  }).join('');

  APP.clusters.forEach((c) => {
    const card = wrap.querySelector('.ccard[data-id="' + c.id + '"]');
    card.querySelector('.ccard-head').addEventListener('click', (e) => {
      if (e.target.closest('button,select,input')) return;
      card.classList.toggle('open');
    });
    card.querySelectorAll('[data-cfg]').forEach((el) => {
      el.addEventListener('input', () => onCfgInput(c.id, card));
      el.addEventListener('change', () => onCfgInput(c.id, card));
    });
    card.querySelectorAll('[data-seg]').forEach((seg) => {
      seg.querySelectorAll('button:not([disabled])').forEach((b) => {
        b.addEventListener('click', () => {
          seg.querySelectorAll('button').forEach((x) => x.classList.remove('active'));
          b.classList.add('active');
          onCfgInput(c.id, card);
        });
      });
    });
  });
  refreshPreviews();
}

function onCfgInput(cid, card) {
  const cfg = getCfg(cid);
  const val = (s) => { const el = card.querySelector('[data-cfg="' + s + '"]'); return el ? el.value : null; };
  const segVal = (s) => { const b = card.querySelector('[data-seg="' + s + '"] button.active'); return b ? b.dataset.val : null; };
  cfg.basis = segVal('basis') || 'allocated';
  cfg.growth = (parseNum(val('growth')) || 0) / 100;
  cfg.redundancy = segVal('redundancy') || 'n1';
  const plat = val('platform');
  if (plat && plat !== cfg.platform) {
    cfg.platform = plat;
    const p = PLATFORMS[plat];
    if (p && CPUS[p.cpu]) {
      const d = CPUS[p.cpu];
      cfg.sockets = p.sockets; cfg.cpu = p.cpu; cfg.cps = d.cores; cfg.ghz = d.ghz; cfg.ramGB = p.ramGB;
      card.querySelector('[data-cfg="sockets"]').value = p.sockets;
      card.querySelector('[data-cfg="cpu"]').value = p.cpu;
      card.querySelector('[data-cfg="cps"]').value = d.cores;
      card.querySelector('[data-cfg="ghz"]').value = d.ghz.toFixed(1);
      card.querySelector('[data-cfg="ramGB"]').value = p.ramGB;
    } else { cfg.platform = 'Custom'; }
  }
  const cpuSel = val('cpu');
  if (cpuSel && cpuSel !== cfg.cpu) {
    cfg.cpu = cpuSel;
    const d = CPUS[cpuSel];
    if (d) {
      cfg.cps = d.cores; cfg.ghz = d.ghz;
      card.querySelector('[data-cfg="cps"]').value = d.cores;
      card.querySelector('[data-cfg="ghz"]').value = d.ghz.toFixed(1);
    }
  }
  cfg.sockets = parseInt(val('sockets')) || 2;
  cfg.cps = parseInt(val('cps')) || 32;
  cfg.ghz = parseFloat(val('ghz')) || 2.5;
  if (cfg.cpu !== 'Custom' && CPUS[cfg.cpu]) {
    const d = CPUS[cfg.cpu];
    if (d.cores !== cfg.cps || Math.abs(d.ghz - cfg.ghz) > 0.001) {
      cfg.cpu = 'Custom';
      card.querySelector('[data-cfg="cpu"]').value = 'Custom';
    }
  }
  cfg.ghzPerVcpu = parseFloat(val('ghzPerVcpu')) || 0.5;
  cfg.ramGB = parseNum(val('ramGB')) || 1024;
  if (val('platform') !== 'Custom' && PLATFORMS[val('platform')]) {
    const p = PLATFORMS[val('platform')];
    if (p.sockets !== cfg.sockets || p.cpu !== cfg.cpu || p.ramGB !== cfg.ramGB) { cfg.platform = 'Custom'; card.querySelector('[data-cfg="platform"]').value = 'Custom'; }
  }
  cfg.cpuOC = parseFloat(val('cpuOC')) || 4;
  cfg.memOC = parseFloat(val('memOC')) || 1.25;
  cfg.hci = segVal('arch') === 'hci';
  cfg.storageTB = parseNum(val('storageTB')) || 15;
  cfg.storageRepl = segVal('storageRepl') || 'mirror';
  card.querySelectorAll('[data-hci-only]').forEach((el) => { el.hidden = !cfg.hci; });
  refreshPreviews();
  queueAutosave();
}

function refreshPreviews() {
  APP.clusters.forEach((c) => {
    const card = document.querySelector('.ccard[data-id="' + c.id + '"]');
    if (!card) return;
    const cfg = getCfg(c.id);
    const r = sizeCluster(c, cfg);
    const set = (k, v) => { const el = card.querySelector('[data-pv="' + k + '"]'); if (el) el.innerHTML = v; };
    const lbl = (k, v) => { const el = card.querySelector('[data-lb="' + k + '"]'); if (el) el.textContent = v; };
    lbl('growth', Math.round(cfg.growth * 100) + '%');
    lbl('cpuOC', cfg.cpuOC.toFixed(1) + ':1');
    lbl('memOC', cfg.memOC.toFixed(2) + ':1');
    lbl('ghzPerVcpu', (cfg.ghzPerVcpu || 0.5).toFixed(1));
    set('hosts', r.finalHosts);
    const bLabel = { cpu: 'CPU-bound', ghz: 'GHz-bound', memory: 'memory-bound', storage: 'storage-bound', balanced: 'balanced' }[r.binding];
    const bColor = r.binding === 'balanced' ? 'var(--green)' : 'var(--amber)';
    set('binding', '<span style="color:' + bColor + '">●</span> ' + bLabel + ' · ' + (cfg.redundancy === 'n1' ? 'N+1' : cfg.redundancy === 'n2' ? 'N+2' : 'no spares'));
    set('spec', '<strong>Target:</strong> ' + esc(cfg.platform) + ' — ' + r.hostLabel +
      ' · ' + cfg.cpuOC.toFixed(1) + ':1 CPU / ' + cfg.memOC.toFixed(2) + ':1 MEM overcommit' +
      (cfg.hci ? ' · HCI ' + fmt1(cfg.storageTB) + ' TB raw/host' : '') +
      ' → <strong>' + r.finalHosts + ' hosts</strong> (' + r.rawHosts + ' + ' + r.spares + ' spare) · ' + fmtInt(r.totalLic) + ' license cores');
  });
}

/* ================= Step 3: results ================= */
function computeResults() { return APP.clusters.map((c) => ({ c, cfg: getCfg(c.id), r: sizeCluster(c, getCfg(c.id)) })); }

function renderResults() {
  APP.results = computeResults();
  renderPlanTab(APP.results);
  renderLicensingTab(APP.results);
  renderFindingsTab(APP.results);
  renderReportTab();
  switchTab('plan');
}

function switchTab(name) {
  document.querySelectorAll('#tabs button').forEach((b) => b.classList.toggle('active', b.dataset.tab === name));
  ['plan', 'licensing', 'findings', 'report'].forEach((t) => { $('tab-' + t).hidden = t !== name; });
}

function mathStep(n, formula, result, isResult) {
  return '<div class="math-step' + (isResult ? ' result' : '') + '"><div class="ms-num">' + n + '</div>' +
    '<div class="ms-formula">' + formula + '</div><div class="ms-result">' + result + '</div></div>';
}

function renderPlanTab(res) {
  const totHosts = res.reduce((a, x) => a + x.r.finalHosts, 0);
  const totLic = res.reduce((a, x) => a + x.r.totalLic, 0);
  const totPhys = res.reduce((a, x) => a + x.r.totalPhysCores, 0);
  const totGHz = res.reduce((a, x) => a + x.r.finalHosts * x.r.perHostGHz, 0);
  const bom = res.map(({ c, cfg, r }) =>
    '<tr><td><strong>' + esc(c.name) + '</strong></td><td>' + esc(cfg.platform) + '</td>' +
    '<td>' + r.hostLabel + '</td><td class="num">' + r.finalHosts + '</td>' +
    '<td class="num">' + fmtInt(r.totalLic) + '</td></tr>').join('');

  const clusters = res.map(({ c, cfg, r }, idx) => {
    let steps = '', n = 0;
    const basisLbl = r.useActual ? 'actual utilization' : 'allocated';
    steps += mathStep(++n, 'demand (' + basisLbl + ') = ' + fmtInt(r.demandVCpu) + ' vCPU · ' + fmtInt(r.demandMemGB) + ' GB RAM' + (cfg.hci ? ' · ' + fmtTB(r.demandStoTB) + ' storage' : ''), '');
    steps += mathStep(++n, 'growth derate = 1 − ' + Math.round(cfg.growth * 100) + '%', '× ' + r.eff.toFixed(2));
    steps += mathStep(++n, 'per-host CPU = ' + cfg.sockets + ' × ' + cfg.cps + ' × ' + cfg.cpuOC.toFixed(1) + ' × ' + r.eff.toFixed(2), fmt1(r.perHostVCpu * r.eff) + ' vCPU');
    steps += mathStep(++n, 'per-host RAM = ' + fmtInt(cfg.ramGB) + ' × ' + cfg.memOC.toFixed(2) + ' × ' + r.eff.toFixed(2), fmtInt(r.perHostMemGB * r.eff) + ' GB');
    steps += mathStep(++n, 'per-host GHz = ' + cfg.sockets + ' × ' + cfg.cps + ' × ' + (cfg.ghz || 2.5).toFixed(1) + ' × ' + r.eff.toFixed(2), fmt1(r.perHostGHz * r.eff) + ' GHz');
    steps += mathStep(++n, 'cpu hosts = ceil(' + fmtInt(r.demandVCpu) + ' ÷ ' + fmt1(r.perHostVCpu * r.eff) + ')', r.cpuHosts);
    steps += mathStep(++n, 'mem hosts = ceil(' + fmtInt(r.demandMemGB) + ' ÷ ' + fmtInt(r.perHostMemGB * r.eff) + ')', r.memHosts);
    steps += mathStep(++n, 'GHz demand = ' + fmtInt(r.demandVCpu) + ' vCPU × ' + (cfg.ghzPerVcpu || 0.5).toFixed(1) + ' GHz/vCPU', fmt1(r.demandGHz) + ' GHz');
    steps += mathStep(++n, 'GHz hosts = ceil(' + fmt1(r.demandGHz) + ' ÷ ' + fmt1(r.perHostGHz * r.eff) + ')', r.ghzHosts);
    if (cfg.hci) steps += mathStep(++n, 'storage hosts = ceil(' + fmtTB(r.demandStoTB) + ' ÷ (' + fmt1(cfg.storageTB) + ' ÷ ' + (STORAGE_REPL[cfg.storageRepl] ? STORAGE_REPL[cfg.storageRepl].factor : 2) + ' × ' + r.eff.toFixed(2) + '))', r.stoHosts);
    const bLabel = { cpu: 'CPU', ghz: 'GHZ', memory: 'MEMORY', storage: 'STORAGE', balanced: 'BALANCED' }[r.binding];
    steps += mathStep(++n, 'binding constraint', '<span style="color:' + (r.binding === 'balanced' ? 'var(--green)' : 'var(--amber)') + '">' + bLabel + '</span>');
    steps += mathStep(++n, 'redundancy (' + (cfg.redundancy === 'n1' ? 'N+1' : cfg.redundancy === 'n2' ? 'N+2' : 'none') + ')', '+' + r.spares + ' spare' + (r.spares === 1 ? '' : 's'));
    steps += mathStep(++n, '<strong>' + r.finalHosts + ' hosts × ' + r.hostLabel + '</strong>', '<strong>' + fmtInt(r.totalLic) + ' license cores</strong>', true);
    return '<div class="panel"><h3>' + (idx + 1) + '. ' + esc(c.name) +
      ' <span class="sub">' + fmtInt(c.vms) + ' VMs · ' + esc(cfg.platform) + '</span></h3>' + steps +
      '<p class="note">Effective capacity at build: ' + Math.round(r.capCpuUsed * 100) + '% of CPU, ' + Math.round(r.capGhzUsed * 100) + '% of GHz, and ' + Math.round(r.capMemUsed * 100) + '% of RAM committed on day one (before growth).</p></div>';
  }).join('');

  $('tab-plan').innerHTML =
    '<div class="stat-grid">' +
    '<div class="stat"><div class="v blue">' + totHosts + '</div><div class="l">New hosts (total)</div></div>' +
    '<div class="stat"><div class="v purple">' + fmtInt(totLic) + '</div><div class="l">vSphere license cores</div></div>' +
    '<div class="stat"><div class="v">' + fmtInt(totPhys) + '</div><div class="l">Physical cores</div></div>' +
    '<div class="stat"><div class="v">' + fmtInt(totGHz) + '</div><div class="l">Total GHz</div></div>' +
    '<div class="stat"><div class="v green">' + res.length + '</div><div class="l">Clusters sized</div></div></div>' +
    '<div class="panel"><h3>Bill of materials</h3><div class="table-scroll"><table class="data">' +
    '<thead><tr><th>Cluster</th><th>Platform</th><th>Host spec</th><th class="num">Hosts</th><th class="num">License cores</th></tr></thead>' +
    '<tbody>' + bom + '<tr><td><strong>Total</strong></td><td></td><td></td><td class="num"><strong>' + totHosts + '</strong></td><td class="num"><strong>' + fmtInt(totLic) + '</strong></td></tr></tbody></table></div>' +
    '<p class="note">Host counts and specs — your VAR quoting tools turn this into dollars. License cores are vSphere compute only.</p></div>' +
    '<h3 style="margin:2rem 0 1rem">Per-cluster worked math</h3>' + clusters;
}

function renderLicensingTab(res) {
  const cur = res.reduce((a, x) => a + (x.c.curLicenseCores || 0), 0);
  const hasCur = res.some((x) => x.c.curLicenseCores != null);
  const totLic = res.reduce((a, x) => a + x.r.totalLic, 0);
  const totPhantom = res.reduce((a, x) => a + x.r.totalPhantom, 0);
  const delta = totLic - cur;
  const rows = res.map(({ c, r }) => {
    const d = c.curLicenseCores != null ? r.totalLic - c.curLicenseCores : null;
    return '<tr><td><strong>' + esc(c.name) + '</strong></td>' +
      '<td class="num">' + (c.curHosts != null ? c.curHosts : '—') + '</td>' +
      '<td class="num">' + (c.curLicenseCores != null ? fmtInt(c.curLicenseCores) : '—') + '</td>' +
      '<td class="num">' + r.finalHosts + '</td>' +
      '<td class="num">' + fmtInt(r.totalLic) + '</td>' +
      '<td class="num">' + (r.phantomPerHost > 0 ? '<span style="color:var(--red)">' + fmtInt(r.totalPhantom) + '</span>' : '<span style="color:var(--green)">0</span>') + '</td>' +
      '<td class="num">' + (d == null ? '—' : '<span style="color:' + (d <= 0 ? 'var(--green)' : 'var(--red)') + '">' + (d <= 0 ? '−' : '+') + fmtInt(Math.abs(d)) + '</span>') + '</td></tr>';
  }).join('');
  $('tab-licensing').innerHTML =
    '<div class="stat-grid">' +
    '<div class="stat"><div class="v">' + (hasCur ? fmtInt(cur) : '—') + '</div><div class="l">License cores today' + (hasCur ? '' : ' (no host data)') + '</div></div>' +
    '<div class="stat"><div class="v purple">' + fmtInt(totLic) + '</div><div class="l">License cores in refresh</div></div>' +
    '<div class="stat"><div class="v ' + (!hasCur ? '' : delta <= 0 ? 'green' : 'red') + '">' + (hasCur ? (delta <= 0 ? '−' : '+') + fmtInt(Math.abs(delta)) : '—') + '</div><div class="l">' + (!hasCur ? 'Delta unavailable' : delta <= 0 ? 'Cores saved' : 'Extra cores') + '</div></div>' +
    '<div class="stat"><div class="v ' + (totPhantom > 0 ? 'amber' : 'green') + '">' + fmtInt(totPhantom) + '</div><div class="l">Phantom cores in build</div></div></div>' +
    '<div class="panel"><h3>Per-cluster licensing</h3><div class="table-scroll"><table class="data">' +
    '<thead><tr><th>Cluster</th><th class="num">Hosts today</th><th class="num">Cores today</th><th class="num">Hosts new</th><th class="num">Cores new</th><th class="num">Phantom</th><th class="num">Delta</th></tr></thead>' +
    '<tbody>' + rows + '</tbody></table></div>' +
    '<p class="note">vSphere per-core subscription: every socket bills <code>max(cores per socket, 16)</code>. Sub-16-core CPUs create phantom cores — licenses paid for silicon that does not exist.</p></div>' +
    '<div class="callout">⚠️ <strong>Indicative math, not a quote.</strong> Editions, bundles (VVF/VCF), vSAN entitlements, and partner pricing change the dollars. Validate against an official Broadcom quote.</div>';
}

function buildFindings(res) {
  const F = [];
  const totHosts = res.reduce((a, x) => a + x.r.finalHosts, 0);
  const totLic = res.reduce((a, x) => a + x.r.totalLic, 0);
  F.push({ sev: 'info', icon: '📋', title: 'Build totals', body: '<strong>' + totHosts + '</strong> new hosts across <strong>' + res.length + '</strong> cluster' + (res.length > 1 ? 's' : '') + ' · <strong>' + fmtInt(totLic) + '</strong> vSphere license cores. Tune anything on the Configure step — this page recomputes when you come back.' });
  res.forEach(({ c, cfg, r }) => {
    if (r.binding === 'memory') F.push({ sev: 'warn', icon: '🧠', title: esc(c.name) + ' is memory-bound', body: 'RAM demand drives this build: ' + r.memHosts + ' hosts for memory vs ' + r.cpuHosts + ' for CPU. Before adding boxes, price more RAM per host — every extra host also drags in ' + r.licPerHost + ' license cores.' });
    else if (r.binding === 'cpu') F.push({ sev: 'info', icon: '⚙️', title: esc(c.name) + ' is CPU-bound', body: 'vCPU demand sets the host count (' + r.cpuHosts + ' vs ' + r.memHosts + ' for RAM). Denser CPUs or a higher CPU overcommit are the levers here.' });
    else if (r.binding === 'ghz') F.push({ sev: 'warn', icon: '⏱️', title: esc(c.name) + ' is GHz-bound', body: 'Clock demand (' + fmt1(r.demandGHz) + ' GHz at ' + (cfg.ghzPerVcpu || 0.5).toFixed(1) + ' GHz/vCPU) sets the host count, not vCPU count — the ' + cfg.cpuOC.toFixed(1) + ':1 overcommit promises more sustained clock per vCPU than ' + (cfg.ghz || 2.5).toFixed(1) + ' GHz cores deliver. Lower the overcommit, pick a higher-clock SKU, or revisit the GHz/vCPU assumption.' });
    else if (r.binding === 'storage') F.push({ sev: 'warn', icon: '💾', title: esc(c.name) + ' is storage-bound (HCI)', body: 'Storage demand needs ' + r.stoHosts + ' hosts but compute only needs ' + Math.max(r.cpuHosts, r.memHosts) + '. Consider fatter storage per host or RAID-5/6 resilience before adding nodes.' });
    else F.push({ sev: 'info', icon: '⚖️', title: esc(c.name) + ' is balanced', body: 'CPU and RAM land on the same host count (' + r.rawHosts + ') — the spec is well matched to the workload mix.' });
    if (r.phantomPerHost > 0) F.push({ sev: 'warn', icon: '👻', title: 'Phantom cores in the ' + esc(c.name) + ' build', body: 'Sub-16-core CPUs bill ' + r.licPerHost + ' cores/host while the silicon only has ' + (cfg.sockets * cfg.cps) + '. That is ' + r.phantomPerHost + ' phantom cores per host — ' + fmtInt(r.totalPhantom) + ' across the build — deleted free by stepping up to ≥16-core CPUs.' });
    if (c.curHosts && r.finalHosts < c.curHosts) F.push({ sev: 'info', icon: '📦', title: 'Consolidation: ' + c.curHosts + ' → ' + r.finalHosts + ' hosts', body: 'The refresh shrinks ' + esc(c.name) + ' by ' + (c.curHosts - r.finalHosts) + ' boxes while carrying ' + Math.round(cfg.growth * 100) + '% growth headroom.' });
    else if (c.curHosts && r.finalHosts > c.curHosts) F.push({ sev: 'warn', icon: '📈', title: 'Host count grows: ' + c.curHosts + ' → ' + r.finalHosts, body: 'Check the drivers — usually growth headroom, redundancy spares, or a lean host spec. Worth a second look before the customer sees it.' });
    if (c.curLicenseCores != null) {
      const d = r.totalLic - c.curLicenseCores;
      F.push({ sev: d <= 0 ? 'info' : 'warn', icon: '🧮', title: 'License cores: ' + fmtInt(c.curLicenseCores) + ' → ' + fmtInt(r.totalLic), body: d <= 0 ? 'Saves ' + fmtInt(-d) + ' cores vs today — mostly phantom-core deletion.' : 'Adds ' + fmtInt(d) + ' cores vs today. The business case needs to carry it.' });
    }
    if (r.useActual && (c.avgCpuUtil > 0.7 || c.avgMemUtil > 0.7)) F.push({ sev: 'warn', icon: '⏱️', title: 'Sized from live utilization — validate', body: esc(c.name) + ' is running hot at export (' + Math.round(c.avgCpuUtil * 100) + '% CPU / ' + Math.round(c.avgMemUtil * 100) + '% MEM). Utilization is point-in-time; confirm against history before committing.' });
    if (cfg.hci && r.demandStoTB <= 0) F.push({ sev: 'warn', icon: '❓', title: 'No storage demand detected for HCI sizing', body: 'The input reported no used storage for ' + esc(c.name) + '. Get a real number before sizing HCI — otherwise the storage host count is a guess.' });
    if (cfg.redundancy === 'n1') F.push({ sev: 'info', icon: '🛡️', title: 'N+2 option for ' + esc(c.name), body: 'Dual-failure protection would take ' + (r.rawHosts + 2) + ' hosts (+1 vs N+1) — ' + fmtInt((r.rawHosts + 2) * r.licPerHost) + ' license cores. Cheap insurance to offer on the quote.' });
  });
  return F;
}

function renderFindingsTab(res) {
  const F = buildFindings(res);
  $('tab-findings').innerHTML = '<div class="panel"><h3>SE talking points <span class="sub">' + F.length + ' findings</span></h3>' +
    F.map((f) => '<div class="finding ' + f.sev + '"><div class="sev">' + f.icon + '</div><div><strong>' + f.title + '</strong><p>' + f.body + '</p></div></div>').join('') + '</div>';
}

function renderReportTab() {
  $('tab-report').innerHTML = '<div class="panel"><h3>📄 Customer-ready briefing</h3>' +
    '<p class="muted">Generates a standalone HTML report — bill of materials, per-cluster worked math, licensing comparison, findings, and methodology. Self-contained (no external dependencies), safe to email. Scrub customer names from cluster labels first if it leaves your org.</p>' +
    '<div class="toolbar"><button class="btn primary" id="dlReportBtn2">⬇ Download HTML report</button></div></div>';
  $('dlReportBtn2').onclick = downloadReport;
}

/* ================= Report ================= */
function buildReportHTML(res) {
  const date = new Date().toISOString().slice(0, 10);
  const srcLbl = APP.source === 'rvtools' ? 'RVTools export (' + APP.fileName + ')' : APP.source === 'demo' ? 'Demo data (synthetic)' : 'Manual entry';
  const totHosts = res.reduce((a, x) => a + x.r.finalHosts, 0);
  const totLic = res.reduce((a, x) => a + x.r.totalLic, 0);
  const F = buildFindings(res);
  const css = 'body{font-family:-apple-system,"Segoe UI",Inter,Roboto,Helvetica,Arial,sans-serif;margin:0;color:#1a1f28;line-height:1.55}' +
    '.wrap{max-width:960px;margin:0 auto;padding:32px 24px}' +
    'h1{font-size:1.9rem;margin:0 0 4px}h2{font-size:1.35rem;margin:2.2rem 0 .8rem;border-bottom:2px solid #4f8cff;padding-bottom:6px}h3{font-size:1.1rem;margin:1.6rem 0 .6rem}' +
    '.meta{color:#5b6572;font-size:.9rem;margin-bottom:1.5rem}' +
    'table{width:100%;border-collapse:collapse;font-size:.88rem;margin:.8rem 0}' +
    'th{text-align:left;color:#5b6572;font-size:.75rem;text-transform:uppercase;letter-spacing:.04em;padding:8px 10px;border-bottom:2px solid #d5dbe4}' +
    'td{padding:8px 10px;border-bottom:1px solid #e6ebf1;vertical-align:top}' +
    '.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}' +
    '.stat{display:inline-block;background:#f2f5fa;border:1px solid #dbe2ec;border-radius:10px;padding:12px 20px;margin:0 10px 10px 0}' +
    '.stat .v{font-size:1.5rem;font-weight:800}.stat .l{font-size:.8rem;color:#5b6572}' +
    '.finding{border:1px solid #dbe2ec;border-left:4px solid #4f8cff;border-radius:8px;padding:12px 16px;margin-bottom:10px;background:#fafbfe}' +
    '.finding.warn{border-left-color:#f5a623}.finding.info{border-left-color:#4f8cff}' +
    '.mono{font-family:ui-monospace,Menlo,Consolas,monospace;font-size:.85em;background:#f2f5fa;padding:1px 6px;border-radius:5px}' +
    '.disclaimer{background:#fff8ec;border:1px solid #f0d9a8;border-radius:10px;padding:14px 18px;margin-top:2rem;font-size:.9rem}' +
    '@media print{.wrap{padding:0}}';
  const bom = res.map(({ c, cfg, r }) =>
    '<tr><td><strong>' + esc(c.name) + '</strong></td><td>' + esc(cfg.platform) + '</td><td>' + r.hostLabel + '</td><td class="num">' + r.finalHosts + '</td><td class="num">' + fmtInt(r.totalLic) + '</td></tr>').join('');
  const perCluster = res.map(({ c, cfg, r }, idx) => {
    const basisLbl = r.useActual ? 'actual utilization' : 'allocated';
    const rows = [
      ['Demand (' + basisLbl + ')', fmtInt(r.demandVCpu) + ' vCPU · ' + fmtInt(r.demandMemGB) + ' GB RAM' + (cfg.hci ? ' · ' + fmtTB(r.demandStoTB) : '')],
      ['Growth headroom', Math.round(cfg.growth * 100) + '% → derate ×' + r.eff.toFixed(2)],
      ['Per-host usable CPU', fmt1(r.perHostVCpu * r.eff) + ' vCPU (' + cfg.sockets + '×' + cfg.cps + ' × ' + cfg.cpuOC.toFixed(1) + ':1)'],
      ['Per-host usable RAM', fmtInt(r.perHostMemGB * r.eff) + ' GB (' + fmtInt(cfg.ramGB) + ' GB × ' + cfg.memOC.toFixed(2) + ':1)'],
      ['Per-host usable GHz', fmt1(r.perHostGHz * r.eff) + ' GHz (' + cfg.sockets + '×' + cfg.cps + ' × ' + (cfg.ghz || 2.5).toFixed(1) + ' GHz)'],
      ['CPU hosts', 'ceil(' + fmtInt(r.demandVCpu) + ' ÷ ' + fmt1(r.perHostVCpu * r.eff) + ') = ' + r.cpuHosts],
      ['Memory hosts', 'ceil(' + fmtInt(r.demandMemGB) + ' ÷ ' + fmtInt(r.perHostMemGB * r.eff) + ') = ' + r.memHosts],
      ['GHz hosts', 'ceil(' + fmt1(r.demandGHz) + ' ÷ ' + fmt1(r.perHostGHz * r.eff) + ') = ' + r.ghzHosts + ' (' + (cfg.ghzPerVcpu || 0.5).toFixed(1) + ' GHz/vCPU assumed)'],
    ];
    if (cfg.hci) rows.push(['Storage hosts (HCI)', 'ceil(' + fmtTB(r.demandStoTB) + ' ÷ usable/host) = ' + r.stoHosts]);
    rows.push(['Binding constraint', { cpu: 'CPU', ghz: 'GHZ', memory: 'MEMORY', storage: 'STORAGE', balanced: 'BALANCED' }[r.binding]]);
    rows.push(['Redundancy', (cfg.redundancy === 'n1' ? 'N+1' : cfg.redundancy === 'n2' ? 'N+2' : 'None') + ' → +' + r.spares + ' spare(s)']);
    rows.push(['<strong>Build</strong>', '<strong>' + r.finalHosts + ' hosts × ' + r.hostLabel + ' — ' + fmtInt(r.totalLic) + ' license cores</strong>']);
    return '<h3>' + (idx + 1) + '. ' + esc(c.name) + ' <span style="color:#5b6572;font-weight:400;font-size:.85rem">· ' + fmtInt(c.vms) + ' VMs · ' + esc(cfg.platform) + '</span></h3>' +
      '<table><tbody>' + rows.map((x) => '<tr><td style="width:38%;color:#5b6572">' + x[0] + '</td><td>' + x[1] + '</td></tr>').join('') + '</tbody></table>';
  }).join('');
  const licRows = res.map(({ c, r }) =>
    '<tr><td><strong>' + esc(c.name) + '</strong></td><td class="num">' + (c.curLicenseCores != null ? fmtInt(c.curLicenseCores) : '—') + '</td><td class="num">' + fmtInt(r.totalLic) + '</td><td class="num">' + (r.phantomPerHost > 0 ? fmtInt(r.totalPhantom) : '0') + '</td></tr>').join('');
  const findings = F.map((f) => '<div class="finding ' + f.sev + '"><strong>' + f.title + '</strong><br><span style="color:#5b6572">' + f.body + '</span></div>').join('');
  return '<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
    '<title>Server Sizer — Refresh Build Plan (' + date + ')</title><style>' + css + '</style></head><body><div class="wrap">' +
    '<h1>Refresh Build Plan</h1><div class="meta">Generated ' + date + ' · Source: ' + esc(srcLbl) + ' · Server Sizer (client-side sizing tool)</div>' +
    '<div><div class="stat"><div class="v">' + totHosts + '</div><div class="l">New hosts</div></div>' +
    '<div class="stat"><div class="v">' + fmtInt(totLic) + '</div><div class="l">vSphere license cores</div></div>' +
    '<div class="stat"><div class="v">' + res.length + '</div><div class="l">Clusters sized</div></div></div>' +
    '<h2>1. Bill of materials</h2><table><thead><tr><th>Cluster</th><th>Platform</th><th>Host spec</th><th class="num">Hosts</th><th class="num">License cores</th></tr></thead><tbody>' + bom +
    '<tr><td><strong>Total</strong></td><td></td><td></td><td class="num"><strong>' + totHosts + '</strong></td><td class="num"><strong>' + fmtInt(totLic) + '</strong></td></tr></tbody></table>' +
    '<h2>2. Per-cluster sizing</h2>' + perCluster +
    '<h2>3. Licensing</h2><table><thead><tr><th>Cluster</th><th class="num">Cores today</th><th class="num">Cores new</th><th class="num">Phantom</th></tr></thead><tbody>' + licRows + '</tbody></table>' +
    '<p style="color:#5b6572;font-size:.9rem">vSphere per-core subscription: each socket bills <span class="mono">max(cores per socket, 16)</span>. Sub-16-core CPUs create phantom cores — licenses paid for silicon that does not exist.</p>' +
    '<h2>4. Findings</h2>' + findings +
    '<h2>5. Methodology</h2><p style="color:#5b6572;font-size:.9rem">Demand per cluster from allocated resources (or allocated × average utilization when export data is present). Host count = ceil(demand ÷ (per-host capacity × overcommit × (1 − growth))) per resource, taking the maximum across CPU, GHz, memory, and HCI storage, plus N+1/N+2 spares. The GHz dimension compares assumed sustained clock per vCPU against physical host clocks (base clock × cores, no overcommit — overcommit is already expressed in the vCPU ratio). CPU models are Intel Xeon 6 SKUs with vendor-published base clocks; turbo frequencies are ignored for sustained sizing. This is capacity sizing, not performance sizing: it does not model IOPS, latency, or NUMA effects. Utilization figures are point-in-time. Overcommit ratios are planner assumptions, not measurements.</p>' +
    '<div class="disclaimer">⚠️ <strong>Indicative analysis, not a quote.</strong> Editions, bundles (VVF/VCF), vSAN entitlements, and partner pricing affect real licensing cost. Validate all figures against an official Broadcom quote before committing to purchases.</div>' +
    '</div></body></html>';
}

function downloadReport() {
  if (!APP.results) return;
  const html = buildReportHTML(APP.results);
  const blob = new Blob([html], { type: 'text/html' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'server-sizer-build-plan-' + new Date().toISOString().slice(0, 10) + '.html';
  document.body.appendChild(a); a.click();
  setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 4000);
}

/* ================= Changelog ================= */
function renderChangelog() {
  const body = $('changelog-body');
  if (!body) return;
  fetch('CHANGELOG.md', { cache: 'no-store' })
    .then((res) => { if (!res.ok) throw new Error('bad status'); return res.text(); })
    .then((md) => {
      let html = '', inList = false;
      const closeList = () => { if (inList) { html += '</ul>'; inList = false; } };
      for (const line of md.split('\n')) {
        if (line.startsWith('## ')) { closeList(); html += '<h4>' + esc(line.slice(3).trim()) + '</h4>'; }
        else if (line.startsWith('- ')) { if (!inList) { html += '<ul>'; inList = true; } html += '<li>' + esc(line.slice(2).trim()) + '</li>'; }
        else if (line.trim() === '' || line.startsWith('# ')) { closeList(); }
        else { closeList(); html += '<p>' + esc(line.trim()) + '</p>'; }
      }
      closeList();
      body.innerHTML = html;
    })
    .catch(() => { body.innerHTML = "<p class='muted'>Changelog unavailable.</p>"; });
}


/* ================= Projects: save / load / export / import ================= */
const LS_AUTO = 'server-sizer:autosave';
const LS_PROJECTS = 'server-sizer:projects';
const PROJECT_VERSION = 1;

function hasDemand(s) { const st = s || APP; return Array.isArray(st.clusters) && st.clusters.length > 0; }
function serializeState() {
  return {
    clusters: APP.clusters, cfgs: APP.cfgs,
    source: APP.source, fileName: APP.fileName,
  };
}
function projectEnvelope(name, state) {
  return {
    app: 'server-sizer', version: PROJECT_VERSION,
    name: (name || 'Untitled project').slice(0, 60),
    savedAt: new Date().toISOString(),
    state: state || serializeState(),
  };
}
function validProject(d) {
  return !!(d && d.app === 'server-sizer' && d.state &&
    Array.isArray(d.state.clusters) && typeof d.version === 'number' && d.version <= PROJECT_VERSION);
}
function slugify(s) { return String(s || 'project').toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40) || 'project'; }
function fmtTime(iso) {
  try { return new Date(iso).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }); }
  catch (e) { return ''; }
}

let toastT = null;
function showToast(html, ms) {
  const t = $('projToast');
  t.innerHTML = html; t.hidden = false;
  clearTimeout(toastT);
  toastT = setTimeout(() => { t.hidden = true; }, ms || 6000);
}
function updateProjName() { $('projName').textContent = APP.projectName || 'Untitled project'; }

function applyProject(env) {
  const s = env.state || {};
  APP.clusters = Array.isArray(s.clusters) ? s.clusters : [];
  APP.cfgs = (s.cfgs && typeof s.cfgs === 'object') ? s.cfgs : {};
  APP.source = s.source || null; APP.fileName = s.fileName || null; APP.results = null;
  APP.projectName = env.name || 'Untitled project';
  if (!APP.clusters.length) { showToast('That project has no demand data \u2014 nothing to restore.'); return; }
  clearMsgs();
  $('landing').hidden = true; $('wizard').hidden = false;
  renderInventory();
  queueAutosave();
  renderConfig();
  setStep(1);
  updateProjName();
  queueAutosave();
}

/* ---- autosave (this browser only) ---- */
let autosaveT = null;
function queueAutosave() { clearTimeout(autosaveT); autosaveT = setTimeout(autosaveNow, 900); }
function autosaveNow() {
  if (!hasDemand()) return;
  try {
    localStorage.setItem(LS_AUTO, JSON.stringify(projectEnvelope(APP.projectName, serializeState())));
    $('projSaved').textContent = '\u00B7 autosaved ' + fmtTime(new Date().toISOString());
  } catch (e) { /* private mode / quota — non-fatal */ }
}
function clearAutosave() { try { localStorage.removeItem(LS_AUTO); } catch (e) {} }

/* ---- named projects (this browser only) ---- */
function getProjects() { try { return JSON.parse(localStorage.getItem(LS_PROJECTS) || '[]'); } catch (e) { return []; } }
function setProjects(list) { try { localStorage.setItem(LS_PROJECTS, JSON.stringify(list.slice(0, 30))); } catch (e) {} }
function renderProjList() {
  const list = getProjects();
  const box = $('projList');
  if (!list.length) { box.innerHTML = '<p class="muted" style="font-size:.85rem">No saved projects yet — name it above and hit <strong>Save project</strong>.</p>'; return; }
  box.innerHTML = list.map((p) => {
    const meta = ((p.state && p.state.clusters) ? p.state.clusters.length : 0) + ' clusters';
    return '<div class="proj-item"><div><div class="nm">' + esc(p.name || 'Untitled project') + '</div>' +
      '<div class="meta">saved ' + esc(fmtTime(p.savedAt)) + ' \u00B7 ' + esc(meta) + '</div></div>' +
      '<div class="ops"><button class="btn ghost" data-load="' + p.id + '">Load</button>' +
      '<button class="btn danger-ghost" data-delp="' + p.id + '">Delete</button></div></div>';
  }).join('');
  box.querySelectorAll('[data-load]').forEach((b) => { b.onclick = () => {
    const p = getProjects().find((x) => x.id === b.dataset.load);
    if (p && validProject(p)) { $('projPanel').hidden = true; applyProject(p); showToast('Loaded project <strong>' + esc(p.name || '') + '</strong>.'); }
    else showToast('Could not load that project — the saved data looks invalid.');
  }; });
  box.querySelectorAll('[data-delp]').forEach((b) => { b.onclick = () => {
    setProjects(getProjects().filter((x) => x.id !== b.dataset.delp));
    renderProjList();
  }; });
}
function saveNamedProject() {
  if (!hasDemand()) { showToast('Load some demand first — there is nothing to save yet.'); return; }
  const input = $('projNameInput').value.trim();
  const name = (input || APP.projectName || 'Untitled project').slice(0, 60);
  const list = getProjects();
  const env = projectEnvelope(name, serializeState());
  env.id = 'p' + Date.now().toString(36);
  const ix = list.findIndex((p) => (p.name || '') === name);
  if (ix >= 0) { env.id = list[ix].id; list[ix] = env; } else list.unshift(env);
  setProjects(list);
  APP.projectName = name; updateProjName();
  $('projNameInput').value = '';
  renderProjList();
  showToast('Project <strong>' + esc(name) + '</strong> saved in this browser.');
  queueAutosave();
}

/* ---- export / import (.json) ---- */
function exportProject() {
  if (!hasDemand()) { showToast('Load some demand first — there is nothing to export yet.'); return; }
  const env = projectEnvelope(APP.projectName, serializeState());
  const blob = new Blob([JSON.stringify(env, null, 2)], { type: 'application/json' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'server-sizer-' + slugify(env.name) + '.json';
  document.body.appendChild(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
  showToast('Exported <strong>' + esc(a.download) + '</strong> — keep it with the engagement files.');
}
function importProjectFile(file) {
  if (!file) return;
  const rd = new FileReader();
  rd.onload = () => {
    try {
      const d = JSON.parse(rd.result);
      if (!validProject(d)) { showToast('<strong>Not a server-sizer project file.</strong> Pick a JSON exported from this app.'); return; }
      $('projPanel').hidden = true;
      applyProject(d);
      showToast('Imported project <strong>' + esc(d.name || 'Untitled') + '</strong>.');
    } catch (e) { showToast('<strong>Could not read that file.</strong> ' + esc(e.message || '')); }
  };
  rd.readAsText(file);
}

function wireProjects() {
  const toggle = () => {
    const p = $('projPanel');
    p.hidden = !p.hidden;
    if (!p.hidden) {
      $('projNameInput').value = APP.projectName === 'Untitled project' ? '' : APP.projectName;
      renderProjList();
      $('projNameInput').focus();
    }
  };
  $('projBtn').onclick = toggle;
  $('projDoSave').onclick = saveNamedProject;
  $('projNameInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') saveNamedProject(); });
  $('projExportBtn').onclick = exportProject;
  $('projImportBtn').onclick = () => $('projImportFile').click();
  $('projImportFile').addEventListener('change', (e) => { importProjectFile(e.target.files[0]); e.target.value = ''; });
  // restore last session, if any
  try {
    const raw = localStorage.getItem(LS_AUTO);
    if (raw) {
      const d = JSON.parse(raw);
      if (validProject(d) && hasDemand(d.state)) {
        applyProject(d);
        showToast('Restored your last session — <strong>' + esc(d.name || '') + '</strong> &nbsp;·&nbsp; <a id="toastFresh">start fresh</a>', 10000);
        const f = $('toastFresh');
        if (f) f.onclick = () => { clearSession(); $('projToast').hidden = true; };
      }
    }
  } catch (e) { /* corrupted autosave — start clean */ }
}

// Export for node unit tests (guarded — undefined in the browser)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { serializeState, projectEnvelope, validProject, hasDemand };
}

/* ================= Wiring ================= */
function clearSession() {
  APP.clusters = []; APP.cfgs = {}; APP.source = null; APP.fileName = null; APP.results = null;
  clearAutosave(); APP.projectName = 'Untitled project'; updateProjName(); $('projSaved').textContent = '';
  $('inventoryWrap').hidden = true;
  $('manualEditor').hidden = true;
  $('manualBody').innerHTML = '';
  $('fileInput').value = '';
  clearMsgs();
  $('wizard').hidden = true;
  $('landing').hidden = false;
  window.scrollTo({ top: 0 });
}

function wireApp() {
  wireProjects();
  document.querySelector('.cta').addEventListener('click', (e) => { e.preventDefault(); startWizard(); });
  // Nav anchor links (How it works / Sizing math / FAQ) target sections inside
  // #landing. When the wizard is open, #landing is hidden and the browser
  // can't scroll to a hidden target — so exit to the landing first, then jump.
  document.querySelectorAll('.nav-links a[href^="#"]').forEach((a) => {
    a.addEventListener('click', (e) => {
      const href = a.getAttribute('href');
      const target = href.length > 1 && document.querySelector(href);
      if (!target) return; // external links (GitHub) behave normally
      e.preventDefault();
      if ($('landing').hidden) { $('wizard').hidden = true; $('landing').hidden = false; }
      target.scrollIntoView({ behavior: 'smooth', block: 'start' });
      history.replaceState(null, '', href);
    });
  });

  const dz = $('dropzone'), fi = $('fileInput');
  dz.addEventListener('click', () => fi.click());
  dz.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); fi.click(); } });
  ['dragover', 'dragenter'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.add('drag'); }));
  ['dragleave', 'drop'].forEach((ev) => dz.addEventListener(ev, (e) => { e.preventDefault(); dz.classList.remove('drag'); }));
  dz.addEventListener('drop', (e) => { const f = e.dataTransfer.files && e.dataTransfer.files[0]; if (f) handleFile(f); });
  fi.addEventListener('change', () => { if (fi.files[0]) handleFile(fi.files[0]); });

  $('demoBtn').onclick = loadDemo;
  $('manualBtn').onclick = () => { clearMsgs(); const me = $('manualEditor'); me.hidden = !me.hidden; if (!me.hidden) { wireManualEditor(); me.scrollIntoView({ behavior: 'smooth' }); } };
  $('manualAddRow').onclick = () => { $('manualBody').insertAdjacentHTML('beforeend', manualRowHTML()); wireManualEditor(); };
  $('manualApply').onclick = applyManual;

  $('backToStartBtn').onclick = () => { APP.clusters = []; APP.source = null; $('inventoryWrap').hidden = true; clearMsgs(); window.scrollTo({ top: 0, behavior: 'smooth' }); };
  $('toConfigBtn').onclick = () => { if (!APP.clusters.length) return; APP.clusters.forEach((c) => getCfg(c.id)); renderConfig(); setStep(2); };
  $('backToDataBtn').onclick = () => setStep(1);
  $('toResultsBtn').onclick = () => { if (!APP.clusters.length) return; renderResults(); setStep(3); };
  $('backToConfigBtn').onclick = () => setStep(2);

  document.querySelectorAll('#tabs button').forEach((b) => b.addEventListener('click', () => switchTab(b.dataset.tab)));
  $('printBtn').onclick = () => window.print();
  $('dlReportBtn').onclick = downloadReport;
  $('clearBtn').onclick = clearSession;
  renderChangelog();
}

if (typeof document !== 'undefined') document.addEventListener('DOMContentLoaded', wireApp);
