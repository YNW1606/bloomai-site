/* BloomAI Excel custom functions (BLOOMAI.*).
 * Registered via functions.json. No build step.
 *
 * Namespace: the functions.json "id"s map to BLOOMAI.ACCOUNT, BLOOMAI.FLOATINGPNL,
 * BLOOMAI.POSITIONS, BLOOMAI.WINRATE (the "BLOOMAI" prefix comes from the manifest
 * namespace registration / the add-in display name).
 *
 * API key resolution:
 *   1. explicit apiKey argument, else
 *   2. OfficeRuntime.storage "bloomai_apikey" (shared with the task pane's localStorage
 *      is browser-scoped, so we also try OfficeRuntime.storage which the task pane mirrors).
 */

var BLOOMAI_DEFAULT_BASE = "https://bloomai-backend-production.up.railway.app";

function _bloomBaseUrl() {
  // OfficeRuntime.storage is async; custom functions read a cached value set at load.
  return (self.__bloomBaseUrl || BLOOMAI_DEFAULT_BASE).replace(/\/+$/, "");
}

// Prime cached settings from OfficeRuntime.storage when the runtime loads.
(function primeSettings() {
  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.storage) {
      OfficeRuntime.storage.getItems(["bloomai_apikey", "bloomai_baseurl"]).then(function (items) {
        self.__bloomApiKey = items["bloomai_apikey"] || null;
        self.__bloomBaseUrl = items["bloomai_baseurl"] || BLOOMAI_DEFAULT_BASE;
      }).catch(function () {});
    }
  } catch (e) { /* ignore */ }
})();

function _resolveKey(apiKey) {
  if (apiKey && String(apiKey).trim()) return String(apiKey).trim();
  if (self.__bloomApiKey) return self.__bloomApiKey;
  throw new CustomFunctions.Error(
    CustomFunctions.ErrorCode.invalidValue,
    "No BloomAI API key. Pass it as an argument or save it in the task pane."
  );
}

function _get(path, apiKey) {
  var key = _resolveKey(apiKey);
  return fetch(_bloomBaseUrl() + path, {
    method: "GET",
    headers: { "X-API-Key": key, "Accept": "application/json" }
  }).then(function (res) {
    if (!res.ok) {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        "BloomAI HTTP " + res.status
      );
    }
    return res.json();
  });
}

/**
 * Account summary as a matrix.
 * @customfunction
 * @param {string} [apiKey]
 * @returns {any[][]}
 */
function ACCOUNT(apiKey) {
  return _get("/api/v1/account", apiKey).then(function (d) {
    var rows = [["Broker", "Account", "Server", "Balance", "Equity", "Live"]];
    (d.accounts || []).forEach(function (a) {
      rows.push([
        a.broker || "", a.account || "", a.server || "",
        a.balance == null ? "" : a.balance,
        a.equity == null ? "" : a.equity,
        a.live ? "LIVE" : "DEMO"
      ]);
    });
    if (rows.length === 1) rows.push(["", "", "", "", "", ""]);
    return rows;
  });
}

/**
 * Total floating (unrealized) P&L.
 * @customfunction
 * @param {string} [apiKey]
 * @returns {number}
 */
function FLOATINGPNL(apiKey) {
  return _get("/api/v1/account", apiKey).then(function (d) {
    return typeof d.floating_pnl === "number" ? d.floating_pnl : Number(d.floating_pnl || 0);
  });
}

/**
 * Open positions as a matrix.
 * @customfunction
 * @param {string} [apiKey]
 * @returns {any[][]}
 */
function POSITIONS(apiKey) {
  return _get("/api/v1/positions", apiKey).then(function (d) {
    var rows = [["Symbol", "Direction", "Lots", "Entry", "SL", "TP", "P&L", "Open Time"]];
    (d.open_positions || []).forEach(function (p) {
      rows.push([
        p.symbol || "", p.direction || "",
        p.lots == null ? "" : p.lots,
        p.entry == null ? "" : p.entry,
        p.sl == null ? "" : p.sl,
        p.tp == null ? "" : p.tp,
        p.pnl == null ? "" : p.pnl,
        p.open_time || ""
      ]);
    });
    if (rows.length === 1) rows.push(["", "", "", "", "", "", "", ""]);
    return rows;
  });
}

/**
 * 30-day win rate percentage.
 * @customfunction
 * @param {string} [apiKey]
 * @returns {number}
 */
function WINRATE(apiKey) {
  return _get("/api/v1/performance", apiKey).then(function (d) {
    var v = (d.last_30d || {}).win_rate_pct;
    return typeof v === "number" ? v : Number(v || 0);
  });
}

// Register with the custom functions runtime.
if (typeof CustomFunctions !== "undefined" && CustomFunctions.associate) {
  CustomFunctions.associate("ACCOUNT", ACCOUNT);
  CustomFunctions.associate("FLOATINGPNL", FLOATINGPNL);
  CustomFunctions.associate("POSITIONS", POSITIONS);
  CustomFunctions.associate("WINRATE", WINRATE);
}
