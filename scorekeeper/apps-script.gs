/**
 * Scorekeeper backend — paste into Extensions > Apps Script of a Google Sheet,
 * then Deploy > New deployment > Web app (Execute as: Me, Who has access: Anyone).
 *
 * Game state lives as JSON on the "_state" tab (one row per game).
 * Each game also gets a readable tab ("Game <id>") with rounds and totals.
 */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.list) return out_(listGames_());
  return out_(getState_(gid_(p.g)));
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var g = gid_(body.g);
  if (body.op === 'deleteGame') return out_(deleteGame_(g));
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var s = applyOp(getState_(g), body);
    s.v = (s.v || 0) + 1;
    s.updated = new Date().toISOString();
    putState_(g, s);
    writeTab_(g, s);
    return out_(s);
  } finally {
    lock.releaseLock();
  }
}

// ---- same logic as the web page ----
function applyOp(s, o) {
  s.players = s.players || []; s.rounds = s.rounds || [];
  var clean = function (v) { v = Math.round(Number(v)); return isFinite(v) ? Math.max(-99999, Math.min(99999, v)) : 0; };
  var has = function (id) { return s.players.some(function (p) { return p.id === id; }); };
  switch (o.op) {
    case 'addPlayer': {
      var id = String(o.id || '').replace(/[^a-z0-9]/gi, '').slice(0, 12);
      var name = String(o.name || '').trim().slice(0, 30);
      if (id && name && !has(id)) s.players.push({ id: id, name: name });
      break;
    }
    case 'renamePlayer': {
      var n = String(o.name || '').trim().slice(0, 30);
      s.players.forEach(function (p) { if (p.id === o.id && n) p.name = n; });
      break;
    }
    case 'removePlayer':
      s.players = s.players.filter(function (p) { return p.id !== o.id; });
      s.rounds.forEach(function (r) { delete r[o.id]; });
      s.rounds = s.rounds.filter(function (r) { return Object.keys(r).length; });
      break;
    case 'setScores': {
      var i = o.index | 0;
      if (i < 0 || i > s.rounds.length) i = s.rounds.length;
      if (i === s.rounds.length) s.rounds.push({});
      var r = s.rounds[i], sc = o.scores || {};
      Object.keys(sc).forEach(function (k) {
        if (!has(k)) return;
        if (sc[k] === null) delete r[k]; else r[k] = clean(sc[k]);
      });
      if (!Object.keys(r).length) s.rounds.splice(i, 1);
      break;
    }
    case 'closeRound': {
      var rr = s.rounds[o.index | 0];
      if (rr) s.players.forEach(function (p) { if (!(p.id in rr)) rr[p.id] = 0; });
      break;
    }
    case 'deleteRound': s.rounds.splice(o.index | 0, 1); break;
    case 'newGame': s.rounds = []; break;
    case 'clearAll': s.players = []; s.rounds = []; break;
    case 'settings':
      if ('target' in o) s.target = Math.max(0, clean(o.target));
      if (o.mode === 'high' || o.mode === 'low') s.mode = o.mode;
      if (typeof o.name === 'string' && o.name.trim()) s.name = o.name.trim().slice(0, 40);
      break;
    case 'restore':
      if (o.state) {
        s.players = o.state.players || []; s.rounds = o.state.rounds || [];
        s.target = o.state.target || 0; s.mode = o.state.mode === 'low' ? 'low' : 'high';
        if (o.state.name) s.name = String(o.state.name).slice(0, 40);
      }
      break;
  }
  return s;
}

// ---- storage ----
function gid_(g) { return String(g || 'table').replace(/[^a-z0-9]/gi, '').slice(0, 20) || 'table'; }

function stateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName('_state');
  if (!sh) { sh = ss.insertSheet('_state'); sh.appendRow(['game', 'json']); }
  return sh;
}

function findRow_(sh, g) {
  var ids = sh.getRange(1, 1, Math.max(1, sh.getLastRow()), 1).getValues();
  for (var i = 1; i < ids.length; i++) if (ids[i][0] === g) return i + 1;
  return 0;
}

function getState_(g) {
  var sh = stateSheet_(), row = findRow_(sh, g);
  if (row) return JSON.parse(sh.getRange(row, 2).getValue());
  return { players: [], rounds: [], target: 500, mode: 'high', v: 0 };
}

function putState_(g, s) {
  var sh = stateSheet_(), row = findRow_(sh, g), json = JSON.stringify(s);
  if (row) sh.getRange(row, 2).setValue(json); else sh.appendRow([g, json]);
}

function listGames_() {
  var sh = stateSheet_(), n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, 2).getValues().map(function (row) {
    try {
      var s = JSON.parse(row[1]);
      return {
        g: row[0], name: s.name || '', updated: s.updated || '', rounds: (s.rounds || []).length, mode: s.mode,
        players: (s.players || []).map(function (p) {
          return { name: p.name, total: (s.rounds || []).reduce(function (a, r) { return a + (r[p.id] || 0); }, 0) };
        })
      };
    } catch (err) { return null; }
  }).filter(function (x) { return x; }).sort(function (a, b) { return a.updated < b.updated ? 1 : -1; });
}

function deleteGame_(g) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = stateSheet_(), row = findRow_(sh, g);
    if (row) sh.deleteRow(row);
    var tab = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Game ' + g);
    if (tab) SpreadsheetApp.getActiveSpreadsheet().deleteSheet(tab);
    return { deleted: g };
  } finally {
    lock.releaseLock();
  }
}

function writeTab_(g, s) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(), name = 'Game ' + g;
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clear();
  if (!s.players.length) return;
  var rows = [['Round'].concat(s.players.map(function (p) { return p.name; }))];
  s.rounds.forEach(function (r, i) {
    rows.push([i + 1].concat(s.players.map(function (p) { return p.id in r ? r[p.id] : ''; })));
  });
  rows.push(['Total'].concat(s.players.map(function (p) {
    return s.rounds.reduce(function (a, r) { return a + (r[p.id] || 0); }, 0);
  })));
  sh.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sh.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sh.getRange(rows.length, 1, 1, rows[0].length).setFontWeight('bold');
}

function out_(s) {
  return ContentService.createTextOutput(JSON.stringify(s)).setMimeType(ContentService.MimeType.JSON);
}
