# BloomAI for Excel &amp; Google Sheets

Pull your live BloomAI trading data — **account**, **open positions**, and **performance** —
straight into Excel or Google Sheets via the public BloomAI REST API.

This is a **standalone** front-end add-in. It does not touch or include the BloomAI backend;
it only consumes the already-live public API.

---

## What you get

**Excel (Office.js task pane)**
- A dark task pane with buttons: **Insert Account / Insert Positions / Insert Performance**.
  Data is written starting at your selected cell on the active sheet.
- Custom functions (BDP-style):
  - `=BLOOMAI.ACCOUNT()` → spilled table: broker, account, server, balance, equity, live
  - `=BLOOMAI.FLOATINGPNL()` → total unrealized P&L (number)
  - `=BLOOMAI.POSITIONS()` → spilled table of open positions
  - `=BLOOMAI.WINRATE()` → 30-day win rate %
  - Each accepts an optional explicit key: `=BLOOMAI.WINRATE("blm_live_xxx")`

**Google Sheets (Apps Script)**
- Custom functions: `=BLOOMAI_BALANCE()`, `=BLOOMAI_FLOATINGPNL()`, `=BLOOMAI_WINRATE()`
  (optional `(apiKey)` arg)
- A **BloomAI** menu: Set API Key, Pull Account, Pull Positions, Pull Performance.

---

## The API it consumes

- **Base URL:** `https://bloomai-backend-production.up.railway.app`
- **Auth header:** `X-API-Key: <your key>`
- **Endpoints**
  - `GET /api/v1/account` → `{ accounts:[{broker,account,server,balance,equity,live}], floating_pnl }`
  - `GET /api/v1/positions` → `{ open_positions:[{symbol,direction,lots,entry,sl,tp,pnl,open_time}], open_count, exposure_by_symbol }`
  - `GET /api/v1/performance` → `{ last_30d:{closed_trades,win_rate_pct,realized_pnl}, recent_closed:[...] }`

> The REST API is **admin-gated behind the `REST_API` feature flag**. If your calls return
> `403`/`404`, the flag isn't enabled for your account — ask your BloomAI admin to enable it.

---

## 1. Get an API key

1. Sign in to **BloomAI**.
2. Go to **Admin → Platform → API Keys**.
3. Create a key and copy it (you'll only see the full value once).

---

## 2. Excel — sideload the add-in

### Host the files
The manifest points at `https://bloomai.technology/excel-addin/…`. Either:
- **Host it:** upload this folder's contents to `https://bloomai.technology/excel-addin/`
  (so `taskpane.html`, `taskpane.js`, `functions.js`, `functions.json`, `commands.html`,
  and `assets/icon-*.png` are reachable at those URLs), **or**
- **Run locally:** serve this folder over HTTPS on `https://localhost:3000/` (e.g.
  `npx office-addin-dev-certs install` then any HTTPS static server) and replace the
  `https://bloomai.technology/excel-addin` URLs in `manifest.xml` with your local URLs.

> Office add-ins require **HTTPS** for all source/icon URLs. Plain `http://` will not load.

### Load the manifest
**Excel on the web / Microsoft 365 desktop:**
- **Insert → Add-ins → Get Add-ins → My Add-ins → Upload My Add-in**, then choose `manifest.xml`.
  (Older builds: **Insert → Office Add-ins → Upload My Add-in**.)

The **BloomAI** button appears on the **Home** tab. Click it to open the task pane.

### Use it
1. In the task pane, set the **Base URL** (defaults to the Railway URL) and paste your **API Key**, then **Save settings** (stored in `localStorage`).
2. Click **Insert Account / Positions / Performance** — output lands at your selected cell.
3. For custom functions, type `=BLOOMAI.WINRATE()` in any cell. With no argument they use the
   key saved in the task pane (mirrored to `OfficeRuntime.storage`); or pass the key explicitly.

> **Custom-function key note:** custom functions run in a separate runtime from the task pane.
> The task pane saves the key to `localStorage`; the functions runtime reads
> `OfficeRuntime.storage`. For reliability, either pass the key as an argument, or save it once
> via the task pane in the same session. (Sharing is best-effort across runtimes.)

---

## 3. Google Sheets — install the script

1. Open your sheet → **Extensions → Apps Script**.
2. Delete the starter code, paste the contents of `google-sheets/Code.gs`, and **Save**.
3. Reload the sheet. A **BloomAI** menu appears.
4. **BloomAI → Set API Key…**, paste your key (stored per-user via `PropertiesService`).
5. Use:
   - Menu items to pull Account / Positions / Performance at the active cell.
   - Functions: `=BLOOMAI_BALANCE()`, `=BLOOMAI_FLOATINGPNL()`, `=BLOOMAI_WINRATE()`.
   - First run will prompt you to **authorize** the script (needed for `UrlFetchApp`).

---

## File tree

```
excel-addin/
├── manifest.xml            # Office Add-in manifest (taskpane + custom functions + ribbon)
├── taskpane.html           # Dark task-pane UI
├── taskpane.js             # Office.js logic: fetch + Excel.run writes
├── commands.html           # Hidden runtime host for FunctionFile + custom functions
├── functions.json          # Custom-function metadata (BLOOMAI.*)
├── functions.js            # Custom-function implementations
├── assets/
│   └── README.txt          # Drop icon-16/32/80.png here
├── google-sheets/
│   └── Code.gs             # Apps Script: custom functions + BloomAI menu
└── README.md
```

---

## Security notes

- Keys are stored **locally** (browser `localStorage` / Office storage for Excel; per-user
  `PropertiesService` for Sheets). They are sent only to the configured base URL over HTTPS.
- Treat API keys like passwords. Revoke from **Admin → Platform → API Keys** if leaked.
- No build step, no dependencies — all vanilla JS.
