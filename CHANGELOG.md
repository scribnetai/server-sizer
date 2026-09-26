# Changelog

## 2026-09-26
- Added 💾 Projects: named saves in this browser, portable JSON export/import, and automatic session restore (your last session reloads on revisit). Saves capture clusters, per-cluster configs, and tuning — nothing uploaded.
- Launched Server Sizer v1: three-step VMware refresh sizing wizard — load demand, configure each cluster, build plan.
- Load demand from an RVTools .xlsx export, manual cluster entry, or one-click demo data.
- Per-cluster config cards: platform presets (Dell R760, Cisco UCS C240 M7, HPE DL380 Gen11, Nutanix NX-8155N, Supermicro, Custom), growth headroom, CPU/RAM overcommit, N+1/N+2, compute vs HCI storage — live host-count previews with binding-constraint callouts.
- Build plan output: per-cluster worked math, bill of materials, vSphere per-core licensing vs current state (incl. phantom cores), SE talking points, downloadable standalone HTML report.
- Fixed: wizard step navigation always showed step 3 — step-switching bug in the landing CTA flow.
- Fixed: "Generate build plan" button did nothing — the results render function was missing.
- Fixed: header nav links (How it works / Sizing math / FAQ) now work while the wizard is open — exits to landing, then scrolls to the section.
- Added: this changelog section, rendered from CHANGELOG.md.
