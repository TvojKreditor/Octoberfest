/**
 * Oktoberfest Birthday Party · 24.10.2026 — приём анкет гостей в Google Таблицу.
 *
 * Скрипт привязан к таблице (Расширения → Apps Script) и развёрнут как веб-приложение:
 * «Запуск от имени: Я», «У кого есть доступ: Все». Подробно — в README.md.
 *
 * Страница-приглашение шлёт POST (Content-Type: text/plain, mode: no-cors) с JSON:
 *   {
 *     "event": "oktoberfest-2026", "group": "<id анкеты>", "sentAt": "<ISO>",
 *     "entries": [ { "id", "name", "foam", "strong", "wishes", "filledBy", "createdAt", "updatedAt" }, … ],
 *     "removedIds": [ "<id>", … ]
 *   }
 * Каждый человек — отдельная строка. Строка опознаётся по паре «группа + id»: повторная отправка
 * (гость нажал «Изменить ответ») обновляет свои строки, а removedIds удаляет убранных спутников.
 * Чужие строки так не изменить: пара группа + id известна только той анкете, что их создала.
 * Ответ — JSON {ok: true}; при mode: no-cors страница его не читает, он для проверки вручную.
 */

var EVENT = 'oktoberfest-2026';
var SHEET_ANSWERS = 'Ответы';
var SHEET_SUMMARY = 'Сводка';
var HEADERS = ['Время', 'Имя', 'Пенное', 'Для тостов', 'Пожелания', 'Кто заполнил', 'Группа', 'Изменено', 'ID'];
var COL_GROUP = 7;   // G
var COL_ID = 9;      // I
var MAX_ENTRIES = 12;
var ID_RE = /^[A-Za-z0-9_-]{1,64}$/;

// Те же варианты, что на странице, — по ним считается «Сводка».
var FOAM = ['Пиво светлое', 'Пиво тёмное', 'Пиво пшеничное', 'Пиво безалкогольное', 'Сидр', 'Медовуха'];
var STRONG = ['Водка', 'Ром', 'Джин', 'Виски или бренди', 'Коньяк', 'Вино белое', 'Вино красное', 'Только пенное'];

/** Приём анкеты со страницы. */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var data = readPayload_(e);
    if (data.event !== EVENT) throw new Error('Чужое событие: ' + String(data.event).slice(0, 40));
    if (!isId_(data.group)) throw new Error('Нет номера анкеты (group)');
    var group = data.group;
    var entries = readEntries_(data.entries);
    var keep = {};
    entries.forEach(function (p) { keep[p.id] = true; });
    var removedIds = readIds_(data.removedIds).filter(function (id) { return !keep[id]; });
    if (!entries.length && !removedIds.length) throw new Error('Пустая анкета');

    lock.waitLock(20000);
    var sheet = ensureSheets_();
    ensureSpace_(sheet, entries.length);
    var removed = removeRows_(sheet, group, removedIds);

    var now = new Date();
    var rows = rowIndex_(sheet, group);
    var inserted = 0, updated = 0, fresh = [];
    entries.forEach(function (p) {
      var values = [cell_(p.name), cell_(p.foam), cell_(p.strong), cell_(p.wishes), cell_(p.filledBy), group];
      var r = rows[p.id];
      if (r) {
        // «Время» первой отправки не трогаем, отмечаем время изменения.
        sheet.getRange(r, 2, 1, 8).setValues([values.concat([now, p.id])]);
        updated++;
      } else {
        fresh.push([now].concat(values, ['', p.id]));
      }
    });
    if (fresh.length) {
      sheet.getRange(sheet.getLastRow() + 1, 1, fresh.length, HEADERS.length).setValues(fresh);
      inserted = fresh.length;
    }
    SpreadsheetApp.flush();

    return json_({ ok: true, inserted: inserted, updated: updated, removed: removed });
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

/** Проверка адреса из браузера: открыть URL веб-приложения — должно ответить ok: true. */
function doGet() {
  return json_({ ok: true, message: 'Приём анкет работает. Анкеты присылает страница-приглашение.' });
}

/**
 * Запустить один раз из редактора (кнопка «Выполнить»): Google спросит разрешение,
 * скрипт создаст и оформит листы «Ответы» и «Сводка». Повторный запуск пересобирает «Сводку»,
 * ответы гостей не трогает.
 */
function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  formatAnswers_(ensureSheets_());
  var summary = ss.getSheetByName(SHEET_SUMMARY);
  summary.clear();
  buildSummary_(summary);

  // Пустой «Лист1» новой таблицы больше не нужен.
  ss.getSheets().forEach(function (sh) {
    if (/^(Лист|Sheet)\d+$/.test(sh.getName()) && sh.getLastRow() === 0 && ss.getSheets().length > 2) {
      ss.deleteSheet(sh);
    }
  });
  ss.setActiveSheet(ss.getSheetByName(SHEET_ANSWERS));
}

/* ---------- чтение и проверка анкеты ---------- */

function readPayload_(e) {
  var raw = e && e.postData && e.postData.contents;
  if (!raw) throw new Error('Пустой запрос');
  try {
    return JSON.parse(raw) || {};
  } catch (x) {
    throw new Error('Не удалось прочитать анкету');
  }
}

function readEntries_(list) {
  list = Array.isArray(list) ? list : [];
  if (list.length > MAX_ENTRIES) throw new Error('В одной анкете больше ' + MAX_ENTRIES + ' человек');
  var byId = {};
  var order = [];
  list.forEach(function (p, i) {
    p = p || {};
    var person = {
      id: isId_(p.id) ? p.id : '',
      name: text_(p.name, 60),
      foam: text_(p.foam, 60),
      strong: text_(p.strong, 60),
      wishes: text_(p.wishes, 500, true),
      filledBy: text_(p.filledBy, 60)
    };
    if (!person.id) throw new Error('У анкеты №' + (i + 1) + ' нет id');
    if (!person.name || !person.foam || !person.strong) throw new Error('Не заполнена анкета №' + (i + 1));
    if (!byId[person.id]) order.push(person.id);
    byId[person.id] = person;   // повтор id в одной отправке — берём последний
  });
  return order.map(function (id) { return byId[id]; });
}

function readIds_(list) {
  return (Array.isArray(list) ? list : []).filter(isId_).slice(0, 50);
}

// Только строка: ID_RE.test(undefined) проверил бы слово "undefined" и пропустил его.
function isId_(v) {
  return typeof v === 'string' && ID_RE.test(v);
}

function text_(v, max, multiline) {
  var s = String(v == null ? '' : v);
  s = multiline
    ? s.replace(/\r\n?/g, '\n').replace(/[ \t]+/g, ' ').replace(/\n{3,}/g, '\n\n').trim()
    : s.replace(/\s+/g, ' ').trim();
  return s.slice(0, max);
}

// Текст, начинающийся с = + - @, таблица приняла бы за формулу — апостроф делает его просто текстом.
function cell_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- лист «Ответы» ---------- */

function ensureSheets_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var answers = ss.getSheetByName(SHEET_ANSWERS);
  if (!answers) answers = ss.insertSheet(SHEET_ANSWERS, 0);
  if (answers.getLastRow() === 0) formatAnswers_(answers);
  if (!ss.getSheetByName(SHEET_SUMMARY)) buildSummary_(ss.insertSheet(SHEET_SUMMARY));
  return answers;
}

function formatAnswers_(sheet) {
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#DCE8F6').setVerticalAlignment('middle');
  sheet.setFrozenRows(1);
  sheet.getRange('A:A').setNumberFormat('dd.MM.yyyy HH:mm');
  sheet.getRange('H:H').setNumberFormat('dd.MM.yyyy HH:mm');
  sheet.getRange('A:I').setVerticalAlignment('top');
  sheet.getRange('E:E').setWrap(true);
  sheet.getRange('G:I').setFontColor('#8496AA');
  [130, 160, 170, 160, 320, 150, 130, 130, 130].forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
}

// Запас пустых строк снизу: строки удаляются по removedIds, а лист не может остаться без незакреплённых строк.
function ensureSpace_(sheet, adding) {
  var need = sheet.getLastRow() + adding + 1;
  var have = sheet.getMaxRows();
  if (have < need) sheet.insertRowsAfter(have, need - have + 50);
}

// Номера строк этой анкеты: id → строка.
function rowIndex_(sheet, group) {
  var last = sheet.getLastRow();
  var index = {};
  if (last < 2) return index;
  var keys = sheet.getRange(2, COL_GROUP, last - 1, COL_ID - COL_GROUP + 1).getValues();
  keys.forEach(function (k, i) {
    if (String(k[0]) === group) index[String(k[COL_ID - COL_GROUP])] = i + 2;
  });
  return index;
}

function removeRows_(sheet, group, ids) {
  if (!ids.length) return 0;
  var drop = {};
  ids.forEach(function (id) { drop[id] = true; });
  var index = rowIndex_(sheet, group);
  var rows = Object.keys(index).filter(function (id) { return drop[id]; })
    .map(function (id) { return index[id]; })
    .sort(function (a, b) { return b - a; });
  rows.forEach(function (r) { sheet.deleteRow(r); });
  return rows.length;
}

/* ---------- лист «Сводка» ---------- */

// Формулы считают прямо по листу «Ответы», поэтому сводка обновляется сама.
function buildSummary_(sheet) {
  var A = "'" + SHEET_ANSWERS + "'!";
  var rows = [];
  var bold = [];
  var head = function (title) {
    bold.push(rows.length + 1);
    rows.push([title, 'Человек', 'Кто']);
  };
  var drinks = function (col, list) {
    var first = rows.length + 1;
    list.forEach(function (name) {
      var r = rows.length + 1;
      rows.push([
        name,
        '=COUNTIF(' + A + col + ':' + col + ',A' + r + ')',
        '=IFERROR(TEXTJOIN(", ",TRUE,FILTER(' + A + 'B:B,' + A + col + ':' + col + '=A' + r + ')),"")'
      ]);
    });
    var lastRow = rows.length;
    rows.push([
      'Другое',
      '=MAX(COUNTA(' + A + col + ':' + col + ')-1-SUM(B' + first + ':B' + lastRow + '),0)',
      ''
    ]);
  };

  rows.push(['Сводка для закупки · Oktoberfest 24.10.2026', '', '']);
  rows.push(['Гостей всего', '=MAX(COUNTA(' + A + 'B:B)-1,0)', '']);
  rows.push(['', '', '']);
  head('Пенное');
  drinks('C', FOAM);
  rows.push(['', '', '']);
  head('Для тостов');
  drinks('D', STRONG);
  rows.push(['', '', '']);
  bold.push(rows.length + 1);
  rows.push(['Пожелания', '', '']);
  var wishesRow = rows.length + 1;
  rows.push([
    '=IFERROR(FILTER(' + A + 'B:B&": "&' + A + 'E:E,' + A + 'E:E<>"",ROW(' + A + 'E:E)>1),"—")',
    '', ''
  ]);

  // Подписи — одним махом, формулы — по одной через setFormula: так их синтаксис
  // (английские имена функций, запятые) не зависит от языка таблицы.
  var formulas = [];
  var labels = rows.map(function (row, r) {
    return row.map(function (v, c) {
      if (String(v).charAt(0) !== '=') return v;
      formulas.push([r + 1, c + 1, v]);
      return '';
    });
  });
  sheet.getRange(1, 1, labels.length, 3).setValues(labels);
  formulas.forEach(function (f) { sheet.getRange(f[0], f[1]).setFormula(f[2]); });

  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sheet.getRange(2, 1, 1, 2).setFontWeight('bold');
  bold.forEach(function (r) {
    sheet.getRange(r, 1, 1, 3).setFontWeight('bold').setBackground('#DCE8F6');
  });
  sheet.getRange(1, 3, wishesRow - 1, 1).setWrap(true);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 360);
  sheet.setFrozenRows(2);
}
