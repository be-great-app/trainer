/**
 * Бекенд трекера тренувань (Google Apps Script).
 * 1) Вставте цей код у Розширення → Apps Script.
 * 2) Запустіть setup() один раз.
 * 3) Розгорніть як веб-застосунок: Виконувати від мого імені, Доступ: усі.
 */

const NAMES = {
  program: 'Програма',
  log: 'Журнал',
  sessions: 'Сесії',
  settings: 'Налаштування'
};

const HEAD = {
  program: ['week', 'date', 'day_no', 'order', 'superset', 'exercise', 'video', 'warmup',
            'weight_kg', 'reps', 'sets', 'rest', 'notes'],
  log: ['id', 'saved_at', 'week', 'date', 'day_no', 'order', 'exercise', 'set_no',
        'weight_kg', 'reps', 'rir', 'is_warmup', 'comment', 'video_url'],
  sessions: ['id', 'date', 'week', 'day_no', 'started_at', 'finished_at', 'wellbeing_1_5', 'comment'],
  settings: ['key', 'value']
};

/** Створює аркуші з заголовками і токени доступу. Безпечно запускати повторно. */
function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(HEAD).forEach(function (k) {
    let sh = ss.getSheetByName(NAMES[k]);
    if (!sh) sh = ss.insertSheet(NAMES[k]);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  // Журнал і сесії зберігаємо як текст, щоб Таблиця не перетворювала дати і числа.
  ['log', 'sessions'].forEach(function (k) {
    ss.getSheetByName(NAMES[k]).getRange('A:Z').setNumberFormat('@');
  });
  ss.getSheetByName(NAMES.program).getRange('B:B').setNumberFormat('@');

  const st = settings_();
  const sh = ss.getSheetByName(NAMES.settings);
  if (!st.WRITE_TOKEN) sh.appendRow(['WRITE_TOKEN', Utilities.getUuid()]);
  if (!st.READ_TOKEN) sh.appendRow(['READ_TOKEN', Utilities.getUuid()]);
  const now = settings_();
  Logger.log('WRITE_TOKEN (для тебе): ' + now.WRITE_TOKEN);
  Logger.log('READ_TOKEN (для тренера): ' + now.READ_TOKEN);
}

function doGet(e) {
  try {
    const p = (e && e.parameter) || {};
    const role = role_(p.token);
    if (!role) return out_({ ok: false, error: 'Невірний токен' });
    const res = {
      ok: true,
      role: role,
      program: rows_('program').filter(function (r) { return String(r.exercise || '').trim(); }),
      log: rows_('log').filter(function (r) { return r.id; }),
      sessions: rows_('sessions').filter(function (r) { return r.id; })
    };
    if (role === 'write') res.read_token = settings_().READ_TOKEN;
    return out_(res);
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents);
    if (role_(b.token) !== 'write') return out_({ ok: false, error: 'Немає прав на запис' });
    upsert_('log', b.sets || []);
    upsert_('sessions', b.sessions || []);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- допоміжне ---------- */

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(k) {
  const sh = SpreadsheetApp.getActive().getSheetByName(NAMES[k]);
  if (!sh) throw new Error('Немає аркуша «' + NAMES[k] + '». Запусти setup().');
  return sh;
}

function settings_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(NAMES.settings);
  const o = {};
  if (!sh || sh.getLastRow() < 2) return o;
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function (r) {
    if (r[0]) o[String(r[0]).trim()] = String(r[1]).trim();
  });
  return o;
}

function role_(token) {
  if (!token) return null;
  const s = settings_();
  if (token === s.WRITE_TOKEN) return 'write';
  if (token === s.READ_TOKEN) return 'read';
  return null;
}

function rows_(k) {
  const sh = sheet_(k);
  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];
  const h = v[0].map(function (x) { return String(x).trim(); });
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return v.slice(1).map(function (r) {
    const o = {};
    h.forEach(function (name, i) {
      let x = r[i];
      if (x instanceof Date) {
        x = name === 'date' ? Utilities.formatDate(x, tz, 'dd.MM.yyyy') : x.toISOString();
      }
      o[name] = x;
    });
    return o;
  });
}

/** Оновлює рядки за id або додає нові; {id, _delete:true} видаляє рядок. */
function upsert_(k, items) {
  if (!items.length) return;
  const sh = sheet_(k);
  const h = HEAD[k];
  const last = sh.getLastRow();
  const idx = {};
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 1).getValues().forEach(function (r, i) {
      if (r[0] !== '') idx[String(r[0])] = i + 2;
    });
  }
  const add = [];
  const del = [];
  const seen = {};
  items.forEach(function (it) {
    const id = String(it.id || '');
    if (!id || seen[id]) return;
    seen[id] = true;
    if (it._delete) {
      if (idx[id]) del.push(idx[id]);
      return;
    }
    const row = h.map(function (c) {
      const x = it[c];
      if (x === null || x === undefined) return '';
      if (typeof x === 'boolean') return x ? 1 : 0;
      return String(x);
    });
    if (idx[id]) sh.getRange(idx[id], 1, 1, h.length).setValues([row]);
    else add.push(row);
  });
  if (add.length) sh.getRange(sh.getLastRow() + 1, 1, add.length, h.length).setValues(add);
  del.sort(function (a, b) { return b - a; }).forEach(function (r) { sh.deleteRow(r); });
}
