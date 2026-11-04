/**
 * Listopadová výzva – backend (Google Apps Script)
 * Vlož do tabulky: Rozšíření → Apps Script → nahraď obsah tímto kódem.
 * 1) Spusť funkci setup (vytvoří listy).
 * 2) Nasadit → Nové nasazení → Webová aplikace → Spustit jako: Já, Přístup: Kdokoli.
 * PIN: první PIN zadaný v appce se uloží (Vlastnosti projektu → PIN).
 * Reset PINu: spusť funkci resetPin a v appce zadej nový.
 */

function getPin(candidate) {
  const props = PropertiesService.getScriptProperties();
  let pin = props.getProperty('PIN');
  if (!pin && candidate) { props.setProperty('PIN', String(candidate)); pin = String(candidate); }
  return pin;
}

function resetPin() { PropertiesService.getScriptProperties().deleteProperty('PIN'); }

const SAMS = ['Báčis', 'Bláža', 'Dan P.', 'Fíla', 'Honza P.', 'Jáchym', 'Kamča', 'Káťa',
  'Libor', 'Mára P.', 'Martin D.', 'Martin Z.', 'Paťas', 'Pepe', 'Táda', 'Tláča'];

const SH_E = 'Záznamy';
const SH_S = 'Směny';
const SH_P = 'SAMové';
const HEAD_E = ['ID', 'Datum', 'Typ', 'SAM', 'Moje pozice', 'Kontext', 'Co jsem řekl / ocenil', 'Reakce SAMa', 'Zapsáno'];
const TZ = 'Europe/Prague';

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const e = ss.getSheetByName(SH_E) || ss.insertSheet(SH_E);
  if (e.getLastRow() === 0) e.appendRow(HEAD_E);
  e.getRange('A:B').setNumberFormat('@');
  e.setFrozenRows(1);
  styleHead(e, HEAD_E.length);
  e.setColumnWidth(7, 320); e.setColumnWidth(8, 260);

  const s = ss.getSheetByName(SH_S) || ss.insertSheet(SH_S);
  if (s.getLastRow() === 0) s.appendRow(['Datum směny']);
  s.getRange('A:A').setNumberFormat('@');
  s.setFrozenRows(1);
  styleHead(s, 1);

  const p = ss.getSheetByName(SH_P) || ss.insertSheet(SH_P);
  if (p.getLastRow() === 0) {
    p.appendRow(['Jméno', 'Aktivní']);
    SAMS.forEach(n => p.appendRow([n, true]));
    p.getRange(2, 2, SAMS.length, 1).insertCheckboxes();
  }
  p.setFrozenRows(1);
  styleHead(p, 2);

  ['List 1', 'Sheet1', 'List1'].forEach(n => {
    const d = ss.getSheetByName(n);
    if (d && ss.getSheets().length > 1) ss.deleteSheet(d);
  });
}

function styleHead(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#1D1D1D').setFontColor('#C6FF00');
}

function doGet() {
  return out({ ok: true, msg: 'Běží. Appka posílá data přes POST.' });
}

function doPost(e) {
  let b = {};
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out({ ok: false, error: 'Neplatný požadavek' }); }
  if (!b.pin || String(b.pin) !== getPin(b.pin)) return out({ ok: false, error: 'Špatný PIN' });

  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ops = Array.isArray(b.ops) ? b.ops : [];
    ops.forEach(op => {
      if (op.action === 'add') add(op.entry);
      else if (op.action === 'delete') del(op.id);
      else if (op.action === 'shift') shift(op.date, op.on);
    });
    return out(list());
  } finally {
    lock.releaseLock();
  }
}

function txt(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy-MM-dd');
  return String(v);
}

function sheet(name) { return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name); }

function list() {
  const ev = sheet(SH_E).getDataRange().getValues().slice(1);
  const entries = ev.filter(r => r[0] !== '').map(r => ({
    id: txt(r[0]), date: txt(r[1]), type: r[2], sam: r[3], role: r[4],
    ctx: r[5], text: r[6], reaction: r[7]
  }));
  const shifts = sheet(SH_S).getDataRange().getValues().slice(1).map(r => txt(r[0])).filter(Boolean);
  const sams = sheet(SH_P).getDataRange().getValues().slice(1).filter(r => r[0] && r[1] === true).map(r => String(r[0]));
  return { ok: true, entries, shifts, sams };
}

function findRow(sh, value) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const col = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < col.length; i++) if (txt(col[i][0]) === String(value)) return i + 2;
  return -1;
}

function add(x) {
  if (!x || !x.id) return;
  const sh = sheet(SH_E);
  if (findRow(sh, x.id) > 0) return; // už zapsáno (opakované odeslání z offline fronty)
  sh.appendRow([x.id, x.date, x.type, x.sam, x.role, x.ctx, x.text, x.reaction,
    Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm')]);
}

function del(id) {
  const sh = sheet(SH_E);
  const r = findRow(sh, id);
  if (r > 0) sh.deleteRow(r);
}

function shift(date, on) {
  const sh = sheet(SH_S);
  const r = findRow(sh, date);
  if (on && r < 0) sh.appendRow([date]);
  if (!on && r > 0) sh.deleteRow(r);
}

function out(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
