/**
 * Family Christmas Lists - Google Apps Script backend.
 * Paste into a Google Sheet: Extensions > Apps Script. Then Deploy > Manage deployments >
 * edit (pencil) > Version: New version > Deploy, so the web app URL stays the same.
 *   Execute as: Me | Who has access: Anyone
 *
 * Privacy:
 *  - For any item you own, the server never sends you the "bought" column.
 *  - The "bought" column is AES-encrypted in the browser with a family
 *    passcode that is never stored here, so opening the Sheet shows only gibberish.
 */
const SS = SpreadsheetApp.openById("1HLiF_IAmb6K_rWB8iK6Ho9AGKAsSziNeLZfB0FYMP8M");
const H = {
  Members: ['key', 'name', 'color'],
  Items:   ['id', 'owner', 'text', 'note', 'bought', 'created'],
  Config:  ['k', 'v']
};
// Tells the page which optional actions this server supports.
const FEATURES = ['addItems', 'editItem'];

function tab_(name) {
  let s = SS.getSheetByName(name);
  if (!s) {
    s = SS.insertSheet(name);
    s.getRange(1, 1, 1, H[name].length).setValues([H[name]]);
    s.setFrozenRows(1);
  }
  return s;
}
function rows_(name) {
  const v = tab_(name).getDataRange().getValues();
  const h = v.shift();
  return v.map((r, i) => {
    const o = { _row: i + 2 };
    h.forEach((k, j) => o[k] = String(r[j] == null ? '' : r[j]));
    return o;
  });
}
function append_(name, obj) {
  const s = tab_(name), h = H[name];
  const rng = s.getRange(s.getLastRow() + 1, 1, 1, h.length);
  rng.setNumberFormat('@');
  rng.setValues([h.map(k => obj[k] == null ? '' : String(obj[k]))]);
}
function setCell_(name, row, col, val) {
  const c = tab_(name).getRange(row, col);
  c.setNumberFormat('@');
  c.setValue(val);
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() { return out_({ ok: true, msg: 'Christmas list API running', features: FEATURES }); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const req = JSON.parse(e.postData.contents);
    return out_(handle_(req));
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function handle_(r) {
  const members = rows_('Members');
  const byKey = {}; members.forEach(m => byKey[m.key] = m);

  if (r.action === 'join') {
    if (!r.key || !r.name || !/^#[0-9a-f]{6}$/i.test(r.color || '')) return { ok: false, error: 'Bad details' };
    if (byKey[r.key] || members.some(m => m.color === r.color)) return { ok: false, error: 'taken' };
    append_('Members', { key: r.key, name: r.name, color: r.color });
    return { ok: true };
  }

  const check = (rows_('Config').find(c => c.k === 'check') || {}).v || '';

  if (r.action === 'list') {
    const items = rows_('Items').map(i => ({
      id: i.id, owner: i.owner, text: i.text, note: i.note, created: i.created,
      bought: i.owner === r.me ? '' : i.bought   // owner never receives their own bought data
    }));
    return { ok: true, members: members.map(m => ({ key: m.key, name: m.name, color: m.color })), items, check, features: FEATURES };
  }

  if (!byKey[r.me]) return { ok: false, error: 'Unknown user' };

  if (r.action === 'setCheck') {
    if (check) return { ok: false, error: 'already set' };
    append_('Config', { k: 'check', v: r.check });
    return { ok: true };
  }

  if (r.action === 'addItem') {
    append_('Items', { id: r.id, owner: r.me, text: r.text, note: r.note || '', bought: '', created: new Date().toISOString() });
    return { ok: true };
  }

  if (r.action === 'addItems') {
    const list = Array.isArray(r.items) ? r.items.slice(0, 100) : [];
    list.forEach(it => append_('Items', { id: it.id, owner: r.me, text: it.text, note: it.note || '', bought: '', created: new Date().toISOString() }));
    return { ok: true, added: list.length };
  }

  const item = rows_('Items').find(i => i.id === r.id);
  if (!item) return { ok: false, error: 'Item not found' };

  if (r.action === 'delItem') {
    if (item.owner !== r.me) return { ok: false, error: 'Not yours' };
    tab_('Items').deleteRow(item._row);
    return { ok: true };
  }

  // Owner changes the wording, links or note of their own item in place.
  // Keeps its position and anything someone has already ticked as bought.
  if (r.action === 'editItem') {
    if (item.owner !== r.me) return { ok: false, error: 'Not yours' };
    if (!String(r.text || '').trim()) return { ok: false, error: 'Item text is empty' };
    setCell_('Items', item._row, H.Items.indexOf('text') + 1, String(r.text));
    setCell_('Items', item._row, H.Items.indexOf('note') + 1, String(r.note || ''));
    return { ok: true };
  }

  if (r.action === 'setBought') {
    if (item.owner === r.me) return { ok: false, error: 'Nope' };
    if ((item.bought || '') !== (r.expected || '')) return { ok: false, conflict: true };  // someone changed it first
    setCell_('Items', item._row, H.Items.indexOf('bought') + 1, r.next || '');
    return { ok: true };
  }

  return { ok: false, error: 'Unknown action' };
}
