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
  comments: 'Коментарі',
  settings: 'Налаштування'
};

const HEAD = {
  program: ['week', 'date', 'day_no', 'order', 'superset', 'exercise', 'video', 'warmup',
            'weight_kg', 'reps', 'sets', 'rest', 'notes'],
  log: ['id', 'saved_at', 'week', 'date', 'day_no', 'order', 'exercise', 'set_no',
        'weight_kg', 'reps', 'rir', 'is_warmup', 'comment', 'video_url'],
  sessions: ['id', 'date', 'week', 'day_no', 'started_at', 'finished_at', 'wellbeing_1_5', 'comment'],
  comments: ['id', 'saved_at', 'week', 'day_no', 'order', 'exercise', 'text', 'author'],
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
  ['log', 'sessions', 'comments'].forEach(function (k) {
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
    let imported = [];
    try { imported = importInbox_(); } catch (err) { imported = ['помилка імпорту: ' + (err.message || err)]; }
    const res = {
      ok: true,
      role: role,
      program: rows_('program').filter(function (r) { return String(r.exercise || '').trim(); }),
      log: rows_('log').filter(function (r) { return r.id; }),
      sessions: rows_('sessions').filter(function (r) { return r.id; }),
      comments: rows_('comments').filter(function (r) { return r.id; })
    };
    if (imported.length) res.imported = imported;
    try { res.pending = pending_(); } catch (err) { res.pending = []; }
    if (role === 'write') res.read_token = settings_().READ_TOKEN;
    return out_(res);
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  // Програма тижня від тренера або учениці (таблиця, скриншоти чи текст) — обробляється до загального блокування,
  // бо importInbox_ бере блокування сам
  try {
    const b0 = JSON.parse(e.postData.contents);
    if (b0.action === 'program') {
      if (!role_(b0.token)) return out_({ ok: false, error: 'Невірний токен' });
      return out_({ ok: true, report: handleProgram_(b0) });
    }
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  }
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const b = JSON.parse(e.postData.contents);
    const role = role_(b.token);
    if (!role) return out_({ ok: false, error: 'Невірний токен' });
    if (role === 'read') {
      // Тренер може лише залишати коментарі до вправ (після перегляду відео)
      upsert_('comments', (b.comments || []).map(function (c) { return Object.assign({}, c, { author: 'coach' }); }));
      return out_({ ok: true });
    }
    upsert_('log', b.sets || []);
    upsert_('sessions', b.sessions || []);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err.message || err) });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- автоімпорт програми з папки на Диску ---------- */

const INBOX = 'Програми тренера';   // папка на Google Диску, куди Claude кладе CSV тижня

/**
 * Бере нові CSV з папки INBOX і дописує їх у аркуш «Програма».
 * Файл, тиждень якого вже є в «Програмі», пропускається (захист від дублів).
 * Оброблені файли запам'ятовуються, тож кожен імпортується лише раз.
 */
function importInbox_() {
  const it = DriveApp.getFoldersByName(INBOX);
  if (!it.hasNext()) return [];
  const folder = it.next();
  const props = PropertiesService.getScriptProperties();
  const done = JSON.parse(props.getProperty('IMPORTED') || '{}');
  const files = [];
  const fi = folder.getFiles();
  while (fi.hasNext()) {
    const f = fi.next();
    if (!done[f.getId()] && /\.csv$/i.test(f.getName())) files.push(f);
  }
  if (!files.length) return [];
  files.sort(function (a, b) { return a.getName() < b.getName() ? -1 : 1; });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = sheet_('program');
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (x) { return String(x).trim(); });
    const weeks = {};
    rows_('program').forEach(function (r) { if (String(r.exercise || '').trim()) weeks[String(r.week)] = true; });
    const report = [];
    files.forEach(function (f) {
      let rows = Utilities.parseCsv(f.getBlob().getDataAsString('UTF-8'));
      rows = rows.filter(function (r) { return r.join('').trim(); });
      if (rows.length && String(rows[0][0]).trim() === 'week') rows.shift();
      const week = rows.length ? String(rows[0][0]).trim() : '';
      if (!week || weeks[week]) {
        report.push(f.getName() + ': тиждень ' + week + ' вже є, пропущено');
      } else {
        const out = rows.map(function (r) {
          return head.map(function (name) {
            const i = HEAD.program.indexOf(name);
            let x = i < 0 ? '' : String(r[i] === undefined ? '' : r[i]).trim();
            if (/^(week|day_no|order|weight_kg|reps|sets)$/.test(name) && /^-?\d+([.,]\d+)?$/.test(x)) x = Number(x.replace(',', '.'));
            return x;
          });
        });
        const start = sh.getLastRow() + 1;
        sh.getRange(start, 2, out.length, 1).setNumberFormat('@');
        sh.getRange(start, 1, out.length, head.length).setValues(out);
        weeks[week] = true;
        report.push(f.getName() + ': додано ' + out.length + ' рядків (тиждень ' + week + ')');
      }
      done[f.getId()] = new Date().toISOString();
    });
    props.setProperty('IMPORTED', JSON.stringify(done));
    return report;
  } finally {
    lock.releaseLock();
  }
}

/* ---------- програма від тренера прямо з додатка ---------- */

const UPLOADS = 'Від тренера — скриншоти і текст';   // підпапка в INBOX: те, що Claude обробить вручну

function inboxFolder_() { const it = DriveApp.getFoldersByName(INBOX); return it.hasNext() ? it.next() : DriveApp.createFolder(INBOX); }
function uploadsFolder_() { const f = inboxFolder_(), it = f.getFoldersByName(UPLOADS); return it.hasNext() ? it.next() : f.createFolder(UPLOADS); }
function programWeeks_() {
  const w = {};
  rows_('program').forEach(function (r) { if (String(r.exercise || '').trim()) w[String(r.week)] = true; });
  return w;
}

/**
 * kind = 'csv'   — таблиця, вставлена в додатку: кладемо CSV у INBOX і одразу імпортуємо (як файли від Claude);
 * kind = 'image' — скриншоти, kind = 'text' — довільний текст: зберігаються в підпапку UPLOADS, їх обробляє Claude.
 */
function handleProgram_(b) {
  const week = String(b.week || '').trim();
  if (!/^\d+$/.test(week)) throw new Error('Не вказано номер тижня');
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');
  if (b.kind === 'csv') {
    if (programWeeks_()[week]) throw new Error('Тиждень ' + week + ' уже є в програмі');
    inboxFolder_().createFile('programa-tyzhden-' + week + '-z-dodatka-' + stamp + '.csv', String(b.csv || ''), MimeType.CSV);
    return importInbox_();
  }
  const f = uploadsFolder_();
  const note = b.dates ? ' (дати: ' + b.dates + ')' : '';
  if (b.kind === 'image') {
    (b.images || []).forEach(function (im, i) {
      f.createFile(Utilities.newBlob(Utilities.base64Decode(im.data), im.mime || 'image/jpeg',
        'trener-tyzhden-' + week + '-' + stamp + '-' + (i + 1) + '.jpg')).setDescription('Тиждень ' + week + note);
    });
    return ['Скриншоти збережено (' + (b.images || []).length + ') — чекають обробки'];
  }
  if (b.kind === 'text') {
    f.createFile('trener-tyzhden-' + week + '-' + stamp + '.txt', 'Тиждень ' + week + note + '\n\n' + String(b.text || ''), MimeType.PLAIN_TEXT);
    return ['Текст збережено — чекає обробки'];
  }
  throw new Error('Невідомий тип програми');
}

/** Надіслані скриншоти/тексти, чий тиждень ще не з'явився в програмі. */
function pending_() {
  const it = DriveApp.getFoldersByName(INBOX);
  if (!it.hasNext()) return [];
  const sub = it.next().getFoldersByName(UPLOADS);
  if (!sub.hasNext()) return [];
  const weeks = programWeeks_(), out = {}, fi = sub.next().getFiles();
  while (fi.hasNext()) {
    const f = fi.next(), m = f.getName().match(/tyzhden-(\d+)/);
    if (!m || weeks[m[1]]) continue;
    const k = m[1];
    if (!out[k]) out[k] = { week: +k, files: 0, kind: '', created: f.getDateCreated().toISOString() };
    out[k].files++;
    out[k].kind = /\.txt$/i.test(f.getName()) ? (out[k].kind === 'image' ? 'mixed' : 'text') : (out[k].kind === 'text' ? 'mixed' : 'image');
  }
  return Object.keys(out).map(function (k) { return out[k]; });
}

/** Запусти вручну один раз після вставки коду — щоб надати доступ до Диска і перевірити імпорт. */
function testImport() {
  Logger.log(importInbox_().join('\n') || 'Нових файлів немає');
}

/* ---------- допоміжне ---------- */

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_(k) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(NAMES[k]);
  if (!sh && k === 'comments') {   // аркуш коментарів тренера створюється сам при першому зверненні
    sh = ss.insertSheet(NAMES[k]);
    sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange('A:Z').setNumberFormat('@');
  }
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
