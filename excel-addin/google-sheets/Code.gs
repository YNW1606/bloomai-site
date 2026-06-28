/**
 * BloomAI — Google Sheets integration (Apps Script).
 *
 * Provides:
 *   - Custom functions: BLOOMAI_BALANCE(apiKey), BLOOMAI_FLOATINGPNL(apiKey), BLOOMAI_WINRATE(apiKey)
 *   - A "BloomAI" menu to pull open positions / account / performance into the sheet.
 *
 * API: https://bloomai-backend-production.up.railway.app
 * Auth: header  X-API-Key: <key>   (create keys in BloomAI Admin -> Platform -> API Keys)
 *
 * Tip: store your key once via the menu (BloomAI -> Set API Key) so the custom
 * functions can be called without passing the key in every cell:  =BLOOMAI_BALANCE()
 */

var BLOOMAI_BASE_URL = 'https://bloomai-backend-production.up.railway.app';
var BLOOMAI_KEY_PROP = 'BLOOMAI_API_KEY';

/* ------------------------------------------------------------------ helpers */

function bloomai_resolveKey_(apiKey) {
  if (apiKey && String(apiKey).trim()) return String(apiKey).trim();
  var saved = PropertiesService.getUserProperties().getProperty(BLOOMAI_KEY_PROP);
  if (saved) return saved;
  throw new Error('No BloomAI API key. Pass it as an argument or use BloomAI menu -> Set API Key.');
}

function bloomai_get_(path, apiKey) {
  var key = bloomai_resolveKey_(apiKey);
  var res = UrlFetchApp.fetch(BLOOMAI_BASE_URL + path, {
    method: 'get',
    headers: { 'X-API-Key': key, 'Accept': 'application/json' },
    muteHttpExceptions: true
  });
  var code = res.getResponseCode();
  var body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('BloomAI HTTP ' + code + ': ' + body);
  }
  return JSON.parse(body);
}

/* ----------------------------------------------------------- custom functions */

/**
 * Total balance across all BloomAI accounts.
 * @param {string=} apiKey Optional API key (else uses stored key).
 * @return {number} Sum of account balances.
 * @customfunction
 */
function BLOOMAI_BALANCE(apiKey) {
  var d = bloomai_get_('/api/v1/account', apiKey);
  var total = 0;
  (d.accounts || []).forEach(function (a) { total += Number(a.balance || 0); });
  return total;
}

/**
 * Total floating (unrealized) P&L.
 * @param {string=} apiKey Optional API key (else uses stored key).
 * @return {number}
 * @customfunction
 */
function BLOOMAI_FLOATINGPNL(apiKey) {
  var d = bloomai_get_('/api/v1/account', apiKey);
  return Number(d.floating_pnl || 0);
}

/**
 * 30-day win rate percentage.
 * @param {string=} apiKey Optional API key (else uses stored key).
 * @return {number}
 * @customfunction
 */
function BLOOMAI_WINRATE(apiKey) {
  var d = bloomai_get_('/api/v1/performance', apiKey);
  return Number((d.last_30d || {}).win_rate_pct || 0);
}

/* ------------------------------------------------------------------- menu UI */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('BloomAI')
    .addItem('Set API Key…', 'bloomai_setApiKey')
    .addSeparator()
    .addItem('Pull Account', 'bloomai_pullAccount')
    .addItem('Pull Positions', 'bloomai_pullPositions')
    .addItem('Pull Performance', 'bloomai_pullPerformance')
    .addToUi();
}

function bloomai_setApiKey() {
  var ui = SpreadsheetApp.getUi();
  var resp = ui.prompt('BloomAI API Key',
    'Paste your API key (Admin -> Platform -> API Keys):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() === ui.Button.OK) {
    var key = resp.getResponseText().trim();
    if (key) {
      PropertiesService.getUserProperties().setProperty(BLOOMAI_KEY_PROP, key);
      ui.alert('BloomAI API key saved.');
    }
  }
}

function bloomai_writeAtSelection_(rows) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var cell = sheet.getActiveCell();
  // normalize to rectangular
  var width = 0;
  rows.forEach(function (r) { if (r.length > width) width = r.length; });
  var grid = rows.map(function (r) {
    var c = r.slice();
    while (c.length < width) c.push('');
    return c;
  });
  sheet.getRange(cell.getRow(), cell.getColumn(), grid.length, width).setValues(grid);
}

function bloomai_pullAccount() {
  try {
    var d = bloomai_get_('/api/v1/account');
    var rows = [['BloomAI — Account', '']];
    rows.push(['Floating P&L', d.floating_pnl]);
    rows.push(['', '']);
    rows.push(['Broker', 'Account', 'Server', 'Balance', 'Equity', 'Live']);
    (d.accounts || []).forEach(function (a) {
      rows.push([a.broker, a.account, a.server, a.balance, a.equity, a.live ? 'LIVE' : 'DEMO']);
    });
    bloomai_writeAtSelection_(rows);
  } catch (e) { SpreadsheetApp.getUi().alert(e.message); }
}

function bloomai_pullPositions() {
  try {
    var d = bloomai_get_('/api/v1/positions');
    var rows = [['BloomAI — Open Positions', '']];
    rows.push(['Open count', d.open_count]);
    rows.push(['', '']);
    rows.push(['Symbol', 'Direction', 'Lots', 'Entry', 'SL', 'TP', 'P&L', 'Open Time']);
    (d.open_positions || []).forEach(function (p) {
      rows.push([p.symbol, p.direction, p.lots, p.entry, p.sl, p.tp, p.pnl, p.open_time]);
    });
    var exp = d.exposure_by_symbol || {};
    var keys = Object.keys(exp);
    if (keys.length) {
      rows.push(['', '']);
      rows.push(['Exposure by symbol', '']);
      keys.forEach(function (k) { rows.push([k, exp[k]]); });
    }
    bloomai_writeAtSelection_(rows);
  } catch (e) { SpreadsheetApp.getUi().alert(e.message); }
}

function bloomai_pullPerformance() {
  try {
    var d = bloomai_get_('/api/v1/performance');
    var d30 = d.last_30d || {};
    var rows = [['BloomAI — Performance (30d)', '']];
    rows.push(['Closed trades', d30.closed_trades]);
    rows.push(['Win rate %', d30.win_rate_pct]);
    rows.push(['Realized P&L', d30.realized_pnl]);
    var recent = d.recent_closed || [];
    if (recent.length) {
      rows.push(['', '']);
      rows.push(['Recent closed trades', '']);
      var keys = Object.keys(recent[0]);
      rows.push(keys);
      recent.forEach(function (t) {
        rows.push(keys.map(function (k) { return t[k]; }));
      });
    }
    bloomai_writeAtSelection_(rows);
  } catch (e) { SpreadsheetApp.getUi().alert(e.message); }
}
