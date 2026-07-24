# HYBRID HOURS LOG

| Field | Value |
|--------|--------|
| **Log name** | HYBRID HOURS |
| **Project** | weighbridge-data-entry (AIS LAN Dashboard) |
| **Repo** | https://github.com/jonocarrillo/weighbridge-data-entry |
| **Host** | Ubuntu Linux (America/Los_Angeles) |
| **Log type** | Work / hours log (for employer presentation) |
| **Opened** | 2026-07-23 (local) |
| **Last entry** | 2026-07-23 21:29 PDT |
| **Time claimed (this day)** | **3.5 hours** (18:00–21:30 PDT) |

---

## TIME SUMMARY — 2026-07-23 (HYBRID HOURS)

**Claimed window:** ~**6:00 PM – 9:30 PM PDT** (**3.5 hours**)  
Includes OS install, machine setup, repo/environment prep, and application work—not only the later coding block.

| Phase | Approx. local time | Duration | Work |
|--------|-------------------|----------|------|
| **A. Prep / Ubuntu install** | **18:00 – 19:30** | **~1.5 hr** | Ubuntu media/prep, install, first boot, desktop initial setup |
| **B. Machine & tooling setup** | **19:30 – 20:50** | **~1.3 hr** | Post-install config, browser/accounts, Grok/agent tooling, Node/runtime, auth/repo access, environment readiness |
| **C. Application session** | **20:50 – 21:30** | **~0.7 hr** | Clone weighbridge-data-entry, run LAN dashboard, ticket UX, Today-Review date fix, editable KPIs, HYBRID HOURS log, nested-folder cleanup |

### Machine-backed anchors (for audit)
| Evidence | Timestamp (PDT) |
|----------|-----------------|
| Installer logs under `/var/log/installer` | ~19:03 – 19:31 |
| First boot / user desktop dirs | ~19:33 |
| GNOME initial setup marked done | ~19:34 |
| Second boot of evening | ~19:44 |
| Repo clone on disk | ~20:50 |
| App edits (ticket/server/dashboard) | ~21:14 – 21:21 |
| HYBRID HOURS log written | ~21:25 – 21:26 |

*Note: Activity before ~19:00 is claimed from operator recollection (ISO/USB/download/prep). From ~19:00 onward, timestamps above support install + setup + app work.*

### Timesheet line (copy/paste)
> **2026-07-23 — 3.5 hr — HYBRID HOURS:** Ubuntu workstation install & setup; AIS weighbridge LAN dashboard clone/deploy; ticket UX (optional truck, pick lists, rate); Today-Review timezone fix; editable KPI cards; work log.

---

## ENTRY 2026-07-23T21:25:00-07:00 — Session summary (application detail)

**Operator:** jonocarrillo / local agent assist  
**Environment:** Ubuntu 26.04 (fresh install this evening), Grok Build, Node v22.17.0 (user-local), dashboard on `0.0.0.0:5000`

### 1. Repository bootstrap
- Cloned / reconstructed private-then-public repo `jonocarrillo/weighbridge-data-entry`.
- Working tree: `/home/annanas/weighbridge-data-entry`
- Runtime: zero-dependency Node + `node:sqlite` (`server.js`, `db.js`).

### 2. LAN service
- Started AIS LAN dashboard: **http://0.0.0.0:5000** (LAN example: `http://192.168.4.204:5000`).
- Database: `data/quickbooks.db`
- Auth token gate: open (no `AIS_DASHBOARD_TOKEN`).

### 3. Tooling notes
- VS Code user-local install attempted; sandbox/`chrome-sandbox` setuid blocked without root; later **uninstalled** per operator request.
- Node installed under `~/.local/node` (no system package).

### 4. Product / UX changes (ticket form)
| Change | Detail |
|--------|--------|
| Truck # optional | Removed required validation / asterisk on ticket save |
| Clickable lists | Product, customer, city — chip pick lists + datalist |
| Rate display | Rate ($/ton) = **amount ÷ net tons**; expected line shows rate × qty |
| Product dropdown source | Merges ticket history + pricing rules (`/api/dropdowns`) |

**Files:** `views/ticket.html`, `server.js`

### 5. Today-Review vs tickets database (bugfix)
**Root cause:** Calendar date mismatch (UTC vs local Pacific).
- Ticket dates often saved via UTC (`toISOString` / `valueAsDate = new Date()`).
- “Today — Review” filtered on **local** calendar day.
- After ~5pm PDT, “today” UTC is already tomorrow → review showed empty while tickets existed.

**Fix:** Local `yyyy-MM-dd` helpers on dashboard, ticket, billing, email; avoid UTC date footguns.  
**Data repair:** Sample ticket `117965` date corrected `2026-07-24` → `2026-07-23` (local day).

**Files:** `views/index.html`, `views/ticket.html`, `views/billing.html`, `views/email.html`

### 6. KPI cards — editable like My Trackers
| Capability | Status |
|------------|--------|
| Edit (✎) | Same control model as trackers |
| Delete (✕) | Same |
| Drag reorder | Within KPI strip |
| Add | **+ Add KPI** or Add Card → style **KPI strip (top row)** |
| Persistence | `/api/dashboard/cards` shared LAN config |

**Model:** Single card list; `style: "kpi"` → top strip; other styles → My Trackers.  
**Defaults:** Total Amount, AR Total, Total Tons, Tickets, Customers (header-filter period).  
**Metrics:** amount, ar, tons, tickets/runs, customers, cash, card, revenue.

**Files:** `views/index.html`, `server.js` (dashboard config merge/preserve on save)

### 7. Verified state (post-fix)
- `/api/day` local today and ticket list aligned after date fix.
- KPI strip and trackers load from shared dashboard card config.
- Dashboard HTTP 200 on port 5000.

### 8. Operator notes
- Nested copy `weighbridge-data-entry/weighbridge-data-entry/` may exist; primary app paths are top-level `server.js` + `views/`.
- Hard-refresh browser (**Ctrl+Shift+R**) after UI deploys.

### 9. Status
**HYBRID HOURS — LOGGED**  
Session changes are recorded. Dashboard LAN service may still be running separately.

---

## Append protocol
Add new entries **above** older detail blocks or under a new `## ENTRY <ISO-8601>` heading with:

```
## ENTRY <timestamp>
**Actor:** …
**Action:** …
**Result:** …
**Files:** …
```

— end of HYBRID HOURS entry —
