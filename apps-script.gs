/**
 * Oktoberfest Birthday Party · 24.10.2026 — приём анкет гостей в Google Таблицу.
 *
 * Скрипт развёрнут как веб-приложение: «Запуск от имени: Я», «У кого есть доступ: Все».
 * Работает и привязанным к таблице (Расширения → Apps Script), и отдельным проектом —
 * тогда в SHEET_ID записан ID таблицы. Подробно — в README.md.
 *
 * Страница-приглашение шлёт POST (Content-Type: text/plain) с JSON:
 *   {
 *     "event": "oktoberfest-2026", "group": "<id анкеты>", "sentAt": "<ISO>",
 *     "entries": [ {
 *        "id", "name", "attend", "costume",
 *        "beer": [..], "mugs", "cider", "contests", "beerFav",
 *        "strong": [..], "how": [..], "pace", "strongFav", "mixers": [..],
 *        "food": [..], "allergy", "music": [..], "wishes", "greeting",
 *        "filledBy", "createdAt", "updatedAt" }, … ],
 *     "removedIds": [ "<id>", … ]
 *   }
 * Каждый человек — отдельная строка. Строка опознаётся по паре «группа + id»: повторная отправка
 * (гость нажал «Изменить ответ») обновляет свои строки, а removedIds удаляет убранных спутников.
 * Чужие строки так не изменить: пара группа + id известна только той анкете, что их создала.
 * Ответ — JSON {ok: true}; страница по нему понимает, что анкета записана.
 */

// ID таблицы из её адреса (…/spreadsheets/d/ЭТО_ID/edit). Пусто — скрипт привязан к таблице и берёт её сам.
var SHEET_ID = '';

var EVENT = 'oktoberfest-2026';
var SHEET_ANSWERS = 'Ответы';
var SHEET_SUMMARY = 'Сводка';
var HEADERS = ['Время', 'Имя', 'Придёт', 'Костюм', 'Пиво и сидр', 'Кружек по 0,5 л', 'Сидр', 'Конкурсы', 'Любимое пенное',
  'Крепкое', 'Как пьёт', 'Темп', 'Любимое крепкое', 'Запивка', 'Еда', 'Аллергия', 'Музыка', 'Пожелания', 'Поздравление',
  'Кто заполнил', 'Группа', 'Изменено', 'ID'];
var COL_GROUP = 21;  // U
var COL_ID = 23;     // W
var MAX_ENTRIES = 12;
var ID_RE = /^[A-Za-z0-9_-]{1,64}$/;
var NO = 'Не смогу';

// Те же варианты, что на странице, — по ним считается «Сводка».
var ATTEND = ['Да, буду!', 'Скорее да', NO];
var COSTUME = ['В баварском костюме', 'Что-нибудь клетчатое', 'Как получится'];
var BEER = ['Светлое пиво', 'Пшеничное пиво', 'Тёмное пиво', 'Нефильтрованное пиво', 'Октоберфестбир', 'Безалкогольное пиво',
  'Сидр', 'Пуаре', 'Медовуха', 'Радлер', 'Белое вино', 'Красное вино', 'Игристое', 'Ничего из этого'];
var CIDER = ['Сладкий', 'Полусладкий', 'Сухой'];
var CONTESTS = ['Я в деле!', 'Посмотрю по настроению', 'Буду болеть за других'];
var STRONG = ['Водка', 'Коньяк', 'Виски', 'Ром', 'Джин', 'Текила', 'Егермейстер', 'Шнапс', 'Настойка', 'Ликёр', 'Крепкое не пью'];
var HOW = ['Шотами', 'Со льдом', 'С колой, тоником, соком', 'В коктейлях'];
var PACE = ['Пара тостов', 'Как пойдёт', 'До утра!'];
var MIXERS = ['Кола', 'Тоник', 'Сок', 'Морс', 'Лимонад', 'Минералка', 'Энергетик', 'Лимон и лайм'];
var FOOD = ['Ем всё', 'Не ем свинину', 'Вегетарианец', 'Без острого', 'Есть аллергия'];
var MUSIC = ['Баварские шлягеры', 'Поп', 'Рок', 'Русские хиты', 'Рэп', 'Электроника', '90-е и нулевые'];

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
      var values = [p.name, p.attend, p.costume, p.beer, p.mugs, p.cider, p.contests, p.beerFav,
        p.strong, p.how, p.pace, p.strongFav, p.mixers, p.food, p.allergy, p.music, p.wishes, p.greeting,
        p.filledBy].map(cell_).concat([group]);
      var r = rows[p.id];
      if (r) {
        // «Время» первой отправки не трогаем, отмечаем время изменения.
        sheet.getRange(r, 2, 1, values.length + 2).setValues([values.concat([now, p.id])]);
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
  var ss = book_();
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

function book_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
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
      attend: ATTEND.indexOf(p.attend) >= 0 ? p.attend : '',
      costume: text_(p.costume, 40),
      beer: list_(p.beer),
      mugs: mugs_(p.mugs),
      cider: text_(p.cider, 40),
      contests: text_(p.contests, 40),
      beerFav: text_(p.beerFav, 80),
      strong: list_(p.strong),
      how: list_(p.how),
      pace: text_(p.pace, 40),
      strongFav: text_(p.strongFav, 80),
      mixers: list_(p.mixers),
      food: list_(p.food),
      allergy: text_(p.allergy, 120),
      music: list_(p.music),
      wishes: text_(p.wishes, 600, true),
      greeting: text_(p.greeting, 600, true),
      filledBy: text_(p.filledBy, 60)
    };
    if (!person.id) throw new Error('У анкеты №' + (i + 1) + ' нет id');
    if (!person.name || !person.attend) throw new Error('Не заполнена анкета №' + (i + 1));
    if (person.attend !== NO && (!person.beer || !person.strong)) throw new Error('Нет ответа про напитки в анкете №' + (i + 1));
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

// Несколько вариантов из списка → одна ячейка «А, Б, В».
function list_(v) {
  var seen = {};
  return (Array.isArray(v) ? v : []).slice(0, 20)
    .map(function (x) { return text_(x, 40); })
    .filter(function (x) { if (!x || seen[x]) return false; seen[x] = true; return true; })
    .join(', ');
}

function mugs_(v) {
  var n = Math.round(Number(v));
  return n >= 1 && n <= 8 ? n : '';
}

// Текст, начинающийся с = + - @, таблица приняла бы за формулу — апостроф делает его просто текстом.
function cell_(s) {
  return typeof s === 'string' && /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- лист «Ответы» ---------- */

function ensureSheets_() {
  var ss = book_();
  var answers = ss.getSheetByName(SHEET_ANSWERS);
  if (!answers) answers = ss.insertSheet(SHEET_ANSWERS, 0);
  if (answers.getLastRow() === 0) formatAnswers_(answers);
  if (!ss.getSheetByName(SHEET_SUMMARY)) buildSummary_(ss.insertSheet(SHEET_SUMMARY));
  return answers;
}

function formatAnswers_(sheet) {
  if (sheet.getMaxColumns() < HEADERS.length) sheet.insertColumnsAfter(sheet.getMaxColumns(), HEADERS.length - sheet.getMaxColumns());
  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#DCE8F6').setVerticalAlignment('middle').setWrap(true);
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
  sheet.getRange('A:A').setNumberFormat('dd.MM.yyyy HH:mm');
  sheet.getRange('V:V').setNumberFormat('dd.MM.yyyy HH:mm');
  sheet.getRange('A:W').setVerticalAlignment('top');
  sheet.getRange('R:S').setWrap(true);
  sheet.getRange('U:W').setFontColor('#8496AA');
  [120, 140, 110, 150, 230, 80, 100, 150, 140, 200, 160, 100, 140, 180, 160, 140, 180, 300, 300, 130, 120, 120, 120]
    .forEach(function (w, i) { sheet.setColumnWidth(i + 1, w); });
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
    rows.push(['', '', '']);
    bold.push(rows.length + 1);
    rows.push([title, 'Человек', 'Кто']);
  };
  var col = function (c) { return A + c + '2:' + c; };
  // Один вариант на человека — точное совпадение.
  var single = function (c, list) {
    list.forEach(function (name) {
      var r = rows.length + 1;
      rows.push([name,
        '=COUNTIF(' + col(c) + ',A' + r + ')',
        '=IFERROR(TEXTJOIN(", ",TRUE,FILTER(' + col('B') + ',' + col(c) + '=A' + r + ')),"")']);
    });
  };
  // Несколько вариантов в одной ячейке через запятую — ищем вхождение.
  var multi = function (c, list) {
    list.forEach(function (name) {
      var r = rows.length + 1;
      rows.push([name,
        '=COUNTIF(' + col(c) + ',"*"&A' + r + '&"*")',
        '=IFERROR(TEXTJOIN(", ",TRUE,FILTER(' + col('B') + ',ISNUMBER(SEARCH(A' + r + ',' + col(c) + ')))),"")']);
    });
  };
  // Свободный текст — всё в одну объединённую ячейку, по строке на человека.
  var merged = [];
  var texts = function (title, c) {
    rows.push(['', '', '']);
    bold.push(rows.length + 1);
    rows.push([title, '', '']);
    merged.push(rows.length + 1);
    rows.push(['=IFERROR(TEXTJOIN(CHAR(10),TRUE,FILTER(' + col('B') + '&": "&' + col(c) + ',' + col(c) + '<>"")),"—")', '', '']);
  };

  rows.push(['Сводка для закупки · Oktoberfest 24.10.2026', '', '']);
  rows.push(['Анкет (человек)', '=COUNTA(' + col('B') + ')', '']);
  rows.push(['Придут точно или скорее', '=COUNTIF(' + col('C') + ',"' + ATTEND[0] + '")+COUNTIF(' + col('C') + ',"' + ATTEND[1] + '")', '']);
  rows.push(['Пива и сидра, литров (по ответам)', '=SUM(' + col('F') + ')*0.5', '']);
  head('Придёт'); single('C', ATTEND);
  head('Пиво, сидр и лёгкое'); multi('E', BEER);
  head('Сидр — какой'); single('G', CIDER);
  head('Пивные конкурсы'); single('H', CONTESTS);
  head('Крепкое'); multi('J', STRONG);
  head('Как пьют крепкое'); multi('K', HOW);
  head('Темп'); single('L', PACE);
  head('Запивка и безалкогольное'); multi('N', MIXERS);
  head('Еда'); multi('O', FOOD);
  head('Музыка'); multi('Q', MUSIC);
  head('Костюм'); single('D', COSTUME);
  texts('Любимое пенное', 'I');
  texts('Любимое крепкое', 'M');
  texts('Аллергии', 'P');
  texts('Пожелания', 'R');
  texts('Поздравления', 'S');

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
  merged.forEach(function (r) { sheet.getRange(r, 1, 1, 3).merge().setWrap(true); });
  formulas.forEach(function (f) { sheet.getRange(f[0], f[1]).setFormula(f[2]); });

  sheet.getRange(1, 1).setFontWeight('bold').setFontSize(13);
  sheet.getRange(2, 1, 3, 2).setFontWeight('bold');
  bold.forEach(function (r) {
    sheet.getRange(r, 1, 1, 3).setFontWeight('bold').setBackground('#DCE8F6');
  });
  sheet.getRange(1, 1, labels.length, 3).setVerticalAlignment('top');
  sheet.getRange(1, 3, labels.length, 1).setWrap(true);
  sheet.setColumnWidth(1, 260);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 380);
  sheet.setFrozenRows(4);
}
