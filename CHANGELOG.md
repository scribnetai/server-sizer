# Changelog

## 2026-09-28
- Fixed: slider track fill was stuck at 50% on the config cards — the theme refresh paints the track with a `--fill` CSS variable, but nothing in the app ever set it. Sliders now update their fill live as you drag, like the other tools.
## 2026-09-28
- TLS certificate provisioned for the `server-sizer.scribnet.io` custom domain (GitHub's stuck DNS check was reset 2026-09-28); HTTPS is now enforced on the site. App-switcher menu links switched from legacy `scribnetai.github.io` URLs to direct `https://<app>.scribnet.io` URLs for all 10 apps (footer/launcher links updated likewise). This entry also covers the net-zero CNAME delete/re-add commits from the DNS-check reset, which carried no changelog entries. Touched: index.html, js/app-switcher.js.


## 2026-09-28
- Migrated legacy `scribnetai.github.io` links to `https://<app>.scribnet.io` for the HTTPS-enforced apps (se-command-center, server-sizer, network-sizer); links to the remaining apps left on the legacy URLs until their TLS certs are issued. Touched: index.html, js/app-switcher.js.

## 2026-09-26
- Restyled to a Physgun-style dark gaming-host vibe: near-black navy background with radial blue glow, Outfit display font, blue→cyan gradients, blue-gradient glyph tiles on section headings, and pill-style segmented option groups.
- Sliders are now custom-styled (glowing blue→cyan fill driven by a --fill var, white thumb with blue halo, min/mid/max scale labels) and every live value shows in a dark-blue pill badge that flashes on change — all outputs update in real time as you drag.
- Each cluster config card now has a live "⚡ What's driving this build" section: animated per-dimension bars (CPU/GHz/RAM/storage) proportional to each dimension's host count, with the binding constraint highlighted — they animate on every input change.
- Build plan results show the same "what's driving this build" bars per cluster, and the big stat numbers (hosts, license cores) now render in gradient text. Sections and cards fade/slide in on scroll.
- Hero gets a banner illustration strip; FAQ items are emoji-prefixed (🔒 💾 📊 🎛️ ⚡ 📄 🧪 🧮 📴 🤔 💵) and section headings carry eyebrow labels.
- Added six more CPUs to the picker: the low-core ROBO tier (6507P 8c @ 3.5 GHz, 6505P 12c @ 2.2, 6517P 16c @ 3.2, 6520P 24c @ 2.4), the new 12-core 6377P entry part (3.1 GHz, 95W), and the 64-core 6710E (2.4 GHz) for smaller efficiency plays. Catalog is now 17 named Xeon SKUs plus Custom.
- Added GHz as a fourth sizing dimension alongside vCPU, RAM, and HCI storage — demand assumes a tunable sustained GHz per vCPU (default 0.5), checked against physical host clocks (cores × base clock, no overcommit). A new "GHz-bound" binding constraint calls out when clock demand sets the host count.
- Added Intel Xeon 6 CPU picker to the per-cluster config: 11 real SKUs (6900P/6700P/6500P P-cores plus the 144-core 6780E) with published base clocks; platform presets now ship with sensible default CPUs. Editing cores or clock flips the CPU to Custom.
- Build plan now shows per-host and total GHz, GHz worked-math steps, GHz in the downloadable report, and a GHz-bound SE finding.
- Fixed the "Is anything stored in my browser?" FAQ — it still claimed nothing goes to localStorage, but the Projects feature (named saves + autosave) does use localStorage on this machine only.
- Added 💾 Projects: named saves in this browser, portable JSON export/import, and automatic session restore (your last session reloads on revisit). Saves capture clusters, per-cluster configs, and tuning — nothing uploaded.
- Launched Server Sizer v1: three-step VMware refresh sizing wizard — load demand, configure each cluster, build plan.
- Load demand from an RVTools .xlsx export, manual cluster entry, or one-click demo data.
- Per-cluster config cards: platform presets (Dell R760, Cisco UCS C240 M7, HPE DL380 Gen11, Nutanix NX-8155N, Supermicro, Custom), growth headroom, CPU/RAM overcommit, N+1/N+2, compute vs HCI storage — live host-count previews with binding-constraint callouts.
- Build plan output: per-cluster worked math, bill of materials, vSphere per-core licensing vs current state (incl. phantom cores), SE talking points, downloadable standalone HTML report.
- Fixed: wizard step navigation always showed step 3 — step-switching bug in the landing CTA flow.
- Fixed: "Generate build plan" button did nothing — the results render function was missing.
- Fixed: header nav links (How it works / Sizing math / FAQ) now work while the wizard is open — exits to landing, then scrolls to the section.
- Added: this changelog section, rendered from CHANGELOG.md.

## 2026-09-27
- Added top-left app-switcher dropdown on the brand mark: one-click jumps to every app in the suite (full index, this page marked).
- Fixed: manual-entry license core math — was undercounting up to 2x; now uses per-socket semantics matching the RVTools path. Renamed manual table header "Cores/host" to "Cores / CPU". Corrected landing privacy banner.
