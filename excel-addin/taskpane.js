/* BloomAI Excel task pane logic.
 * Vanilla Office.js — no build step.
 * Fetches from the BloomAI public REST API and writes 2D arrays into the active worksheet.
 */
(function () {
  "use strict";

  var DEFAULT_BASE = "https://bloomai-backend-production.up.railway.app";
  var LS_KEY = "bloomai_apikey";
  var LS_URL = "bloomai_baseurl";

  function $(id) { return document.getElementById(id); }

  function getBaseUrl() {
    var v = ($("baseUrl").value || "").trim();
    return (v || DEFAULT_BASE).replace(/\/+$/, "");
  }
  function getApiKey() {
    return ($("apiKey").value || "").trim();
  }

  function setStatus(kind, msg) {
    var el = $("status");
    el.className = kind; // ok | err | info
    el.textContent = msg;
  }

  function loadSettings() {
    $("baseUrl").value = localStorage.getItem(LS_URL) || DEFAULT_BASE;
    $("apiKey").value = localStorage.getItem(LS_KEY) || "";
  }

  function saveSettings() {
    localStorage.setItem(LS_URL, getBaseUrl());
    localStorage.setItem(LS_KEY, getApiKey());
    setStatus("ok", "Settings saved locally.");
  }

  // ---- REST helper -----------------------------------------------------------
  function apiGet(path) {
    var key = getApiKey();
    if (!key) {
      return Promise.reject(new Error("No API key set. Paste your key and click Save."));
    }
    var url = getBaseUrl() + path;
    return fetch(url, {
      method: "GET",
      headers: { "X-API-Key": key, "Accept": "application/json" }
    }).then(function (res) {
      if (!res.ok) {
        return res.text().then(function (t) {
          var detail = t;
          try { detail = JSON.parse(t).detail || t; } catch (e) {}
          throw new Error("HTTP " + res.status + ": " + (detail || res.statusText));
        });
      }
      return res.json();
    });
  }

  // ---- formatting helpers ----------------------------------------------------
  function num(v) { return (v === null || v === undefined || v === "") ? "" : v; }

  function accountRows(data) {
    var rows = [["BloomAI — Account", ""]];
    rows.push(["Floating P&L", num(data.floating_pnl)]);
    rows.push(["", ""]);
    rows.push(["Broker", "Account", "Server", "Balance", "Equity", "Live"]);
    (data.accounts || []).forEach(function (a) {
      rows.push([
        num(a.broker), num(a.account), num(a.server),
        num(a.balance), num(a.equity), a.live ? "LIVE" : "DEMO"
      ]);
    });
    return rows;
  }

  function positionsRows(data) {
    var rows = [["BloomAI — Open Positions", ""]];
    rows.push(["Open count", num(data.open_count)]);
    rows.push(["", ""]);
    rows.push(["Symbol", "Direction", "Lots", "Entry", "SL", "TP", "P&L", "Open Time"]);
    (data.open_positions || []).forEach(function (p) {
      rows.push([
        num(p.symbol), num(p.direction), num(p.lots), num(p.entry),
        num(p.sl), num(p.tp), num(p.pnl), num(p.open_time)
      ]);
    });
    var exp = data.exposure_by_symbol || {};
    var expKeys = Object.keys(exp);
    if (expKeys.length) {
      rows.push(["", ""]);
      rows.push(["Exposure by symbol", ""]);
      expKeys.forEach(function (k) { rows.push([k, num(exp[k])]); });
    }
    return rows;
  }

  function performanceRows(data) {
    var d30 = data.last_30d || {};
    var rows = [["BloomAI — Performance (30d)", ""]];
    rows.push(["Closed trades", num(d30.closed_trades)]);
    rows.push(["Win rate %", num(d30.win_rate_pct)]);
    rows.push(["Realized P&L", num(d30.realized_pnl)]);
    var recent = data.recent_closed || [];
    if (recent.length) {
      rows.push(["", ""]);
      rows.push(["Recent closed trades", ""]);
      // header from union of keys (stable order from first row)
      var keys = Object.keys(recent[0]);
      rows.push(keys);
      recent.forEach(function (t) {
        rows.push(keys.map(function (k) { return num(t[k]); }));
      });
    }
    return rows;
  }

  // pad ragged rows to equal width so Excel range matches the 2D array
  function rectangular(rows) {
    var w = 0;
    rows.forEach(function (r) { if (r.length > w) w = r.length; });
    return rows.map(function (r) {
      var copy = r.slice();
      while (copy.length < w) copy.push("");
      return copy;
    });
  }

  // ---- write into Excel ------------------------------------------------------
  function writeRows(rows) {
    var grid = rectangular(rows);
    return Excel.run(function (ctx) {
      var sheet = ctx.workbook.worksheets.getActiveWorksheet();
      var anchor = ctx.workbook.getSelectedRange();
      anchor.load("address, rowIndex, columnIndex");
      return ctx.sync().then(function () {
        var target = sheet.getRangeByIndexes(
          anchor.rowIndex, anchor.columnIndex, grid.length, grid[0].length
        );
        target.values = grid;
        target.format.autofitColumns();
        // bold the title row
        sheet.getRangeByIndexes(anchor.rowIndex, anchor.columnIndex, 1, grid[0].length)
          .format.font.bold = true;
        return ctx.sync();
      });
    });
  }

  function run(label, path, mapper) {
    setStatus("info", "Fetching " + label + "…");
    apiGet(path)
      .then(function (data) { return writeRows(mapper(data)); })
      .then(function () { setStatus("ok", label + " inserted at selection."); })
      .catch(function (err) { setStatus("err", (err && err.message) || String(err)); });
  }

  // ---- wire up ---------------------------------------------------------------
  Office.onReady(function (info) {
    if (info.host !== Office.HostType.Excel) {
      // still allow settings outside Excel, but warn
    }
    loadSettings();

    $("saveBtn").addEventListener("click", saveSettings);
    $("showKey").addEventListener("change", function () {
      $("apiKey").type = this.checked ? "text" : "password";
    });

    $("btnAccount").addEventListener("click", function () {
      run("Account", "/api/v1/account", accountRows);
    });
    $("btnPositions").addEventListener("click", function () {
      run("Positions", "/api/v1/positions", positionsRows);
    });
    $("btnPerformance").addEventListener("click", function () {
      run("Performance", "/api/v1/performance", performanceRows);
    });
  });
})();
