/**
 * ОТДЕЛЬНЫЙ файл для основного проекта (тот, где лист «Задачи Bitrix» и
 * меню «Bitrix» из BitrixTasksToSheets.gs).
 *
 * Кнопка «Учёт времени за период» выгружает на отдельный лист ВЕСЬ ваш учёт
 * времени за произвольный период — по одной строке на КАЖДУЮ запись учёта
 * (дата/время, длительность, комментарий, задача, проект, ссылка).
 *
 * ЧТО ИМЕННО БЕРЁТСЯ:
 *   • записи, где время учитывали ВЫ (USER_ID = вы) — это ваш личный тайм-лог;
 *   • по всем задачам, где вы — исполнитель (RESPONSIBLE_ID = вы).
 *   Хотите включить и задачи, где вы НЕ исполнитель, но там учитывали время, —
 *   поставьте TP_ONLY_MY_TASKS = false ниже.
 *
 * ПЕРИОД задаётся ОДНОЙ строкой в формате  дд.мм.гггг-дд.мм.гггг
 *   пример:  07.07.2026-15.07.2026
 *
 * Файл самостоятельный: все функции с префиксом tp, чтобы не конфликтовать
 * с основным скриптом в том же проекте Apps Script. onOpen тут НЕТ — пункт
 * меню добавляется в основной скрипт (BitrixTasksToSheets.gs), одной строкой:
 *     .addItem('Учёт времени за период', 'tpExportTimeByPeriod')
 *
 * ВЕБХУК — общий с основным скриптом (свойство BITRIX_WEBHOOK). Если вебхук
 * уже сохранён через меню «Bitrix» → «Указать вебхук», повторно вводить не нужно.
 *
 * ЗАПУСК ИЗ РЕДАКТОРА (без меню/диалога): впишите период в TP_PERIOD ниже и
 * запустите функцию tpExportTimeByPeriod. Результат и сообщения — в журнале.
 */

// ───────────────────── Настройки ─────────────────────
var TP_DEST_SHEET     = 'Учёт времени за период'; // лист результата (перезаписывается)
var TP_PROP_WEBHOOK   = 'BITRIX_WEBHOOK';         // тот же ключ, что в основном скрипте
var TP_RESPONSIBLE_ID = '';                       // '' — владелец вебхука; иначе Bitrix ID строкой
var TP_ONLY_MY_TASKS  = true;                     // true — только задачи, где я исполнитель
var TP_PERIOD         = '';                       // из редактора: 'дд.мм.гггг-дд.мм.гггг'
var TP_TIMEZONE       = '';                       // '' — пояс скрипта; иначе напр. 'Europe/Moscow'
var TP_MAX_ELAPSED    = 5000;                     // предохранитель: максимум записей учёта
var TP_PAGE_SIZE      = 50;                       // task.elapseditem.getlist отдаёт по 50

// Заголовки листа — по одной строке на каждую запись учёта времени
var TP_HEADERS = [
  'Дата и время', 'Длительность', 'Длительность, мин', 'Комментарий',
  'Задача', 'ID задачи', 'Проект/Группа', 'Ссылка'
];

// ───────────────────── Главное действие ─────────────────────
function tpExportTimeByPeriod() {
  var ui = tpUi_();
  var tz = TP_TIMEZONE || Session.getScriptTimeZone();

  // 1. Период одной строкой: диалог или TP_PERIOD (запуск из редактора)
  var period = tpPromptPeriod_(ui);
  if (!period) return; // отменили диалог

  // 2. «Я» = RESPONSIBLE_ID, иначе владелец вебхука, иначе user.current
  var me = String(TP_RESPONSIBLE_ID || tpWebhookOwnerId_() ||
                  tpCallBitrix_('user.current', {}).result.ID);

  // 3. Все мои записи учёта времени за период (постранично, один метод)
  var items = tpFetchElapsed_(me, period.fromApi, period.toApi);

  // 4. Разворачиваем в строки. По задаче подтягиваем название/группу/исполнителя
  //    (с кэшем), чтобы не дёргать REST по одной и той же задаче.
  var portalBase = tpGetWebhook_().replace(/\/rest\/.*/, '');
  var taskCache = {}, groupCache = {};
  var rows = [];
  var totalSec = 0, kept = 0, taskSet = {};

  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    var taskId = String(tpPick_(it, ['TASK_ID', 'taskId']));
    if (!taskId || taskId === '0') continue;
    var info = tpTaskInfo_(taskId, taskCache);

    // только задачи, где я исполнитель (если фильтр включён)
    if (TP_ONLY_MY_TASKS && String(info.responsibleId) !== me) continue;

    var seconds = tpSeconds_(it);
    var created = tpPick_(it, ['CREATED_DATE', 'createdDate']);
    var comment = tpText_(tpPick_(it, ['COMMENT_TEXT', 'commentText']));
    totalSec += seconds; kept++; taskSet[taskId] = true;

    rows.push([
      tpFmtDate_(created, tz),
      tpDuration_(seconds),
      tpRound1_(seconds / 60),
      comment,
      info.title,
      Number(taskId) || taskId,
      tpGroupName_(info.groupId, groupCache),
      tpTaskUrl_(portalBase, info.responsibleId || me, taskId)
    ]);
  }

  rows.sort(function (a, b) { return String(b[0]).localeCompare(String(a[0])); }); // позже записанные выше

  // 5. Пишем лист
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TP_DEST_SHEET) || ss.insertSheet(TP_DEST_SHEET);
  sheet.clear();

  var title = 'Учёт времени за ' + period.from + ' — ' + period.to +
    ' · пользователь ID ' + me +
    (TP_ONLY_MY_TASKS ? ' · только мои задачи' : ' · все задачи') +
    ' · записей: ' + kept +
    ' · задач: ' + Object.keys(taskSet).length +
    ' · итого: ' + tpDuration_(totalSec) + ' (' + tpRound1_(totalSec / 60) + ' мин)';
  sheet.getRange(1, 1, 1, TP_HEADERS.length).merge().setValue(title)
    .setFontWeight('bold').setBackground('#e8eaf6');

  var out = [TP_HEADERS].concat(rows);
  if (rows.length === 0) out.push(['(за этот период записей учёта времени нет)']);
  out = tpPadRows_(out);

  var startRow = 2;
  sheet.getRange(startRow, 1, out.length, TP_HEADERS.length).setValues(out);
  sheet.getRange(startRow, 1, 1, TP_HEADERS.length).setFontWeight('bold').setBackground('#f1f8e9');
  sheet.setFrozenRows(startRow);
  sheet.autoResizeColumns(1, TP_HEADERS.length);

  var msg = 'Учёт времени за ' + period.from + '—' + period.to + ': записей ' + kept +
    ', итого ' + tpDuration_(totalSec) + '. Лист «' + TP_DEST_SHEET + '».';
  Logger.log(title + '\n' + msg);
  if (ui) ss.toast(msg, 'Bitrix', 8);
}

// ───────────────────── Период (одной строкой) ─────────────────────
// Спрашивает период ОДНОЙ строкой; из редактора берёт TP_PERIOD.
// Возвращает { from, to, fromApi, toApi } или null, если отменили.
function tpPromptPeriod_(ui) {
  var raw;
  if (ui) {
    var r = ui.prompt('Учёт времени за период',
      'Введите период ОДНОЙ строкой в формате  дд.мм.гггг-дд.мм.гггг\n' +
      'например:  07.07.2026-15.07.2026',
      ui.ButtonSet.OK_CANCEL);
    if (r.getSelectedButton() !== ui.Button.OK) return null;
    raw = r.getResponseText();
  } else {
    raw = TP_PERIOD;
    if (!String(raw || '').trim()) {
      throw new Error('Запуск из редактора: впишите период в TP_PERIOD ' +
        '(формат дд.мм.гггг-дд.мм.гггг, например 07.07.2026-15.07.2026).');
    }
  }
  var period = tpParsePeriod_(raw);
  if (!period) {
    var m = 'Неверный формат периода. Ожидается  дд.мм.гггг-дд.мм.гггг, ' +
            'например 07.07.2026-15.07.2026.';
    if (ui) ui.alert(m);
    throw new Error(m);
  }
  return period;
}

// 'дд.мм.гггг-дд.мм.гггг' → { from, to (для показа), fromApi, toApi } либо null
function tpParsePeriod_(raw) {
  var s = String(raw || '').trim().replace(/[–—]/g, '-'); // en/em-тире → обычный дефис
  var parts = s.split(/\s*-\s*/);
  if (parts.length !== 2) return null;
  var a = tpParseDay_(parts[0]);
  var b = tpParseDay_(parts[1]);
  if (!a || !b) return null;
  if (a.ymd > b.ymd) { var t = a; a = b; b = t; } // перепутали местами — поправим
  return {
    from: a.dmy, to: b.dmy,
    fromApi: a.ymd + ' 00:00:00',
    toApi:   b.ymd + ' 23:59:59'
  };
}

// 'дд.мм.гггг' → { ymd:'гггг-мм-дд', dmy:'дд.мм.гггг' } либо null (с проверкой даты)
function tpParseDay_(str) {
  var m = String(str || '').trim().match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (!m) return null;
  var d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]);
  var dt = new Date(y, mo - 1, d);
  if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) {
    return null; // напр. 31.02.2026 — такой даты нет
  }
  function p(n) { return (n < 10 ? '0' : '') + n; }
  return { ymd: y + '-' + p(mo) + '-' + p(d), dmy: p(d) + '.' + p(mo) + '.' + y };
}

// ───────────────────── Записи учёта времени (постранично) ─────────────────────
// task.elapseditem.getlist: параметры ПОЗИЦИОННЫЕ (порядок ORDER, FILTER, SELECT,
// PARAMS важнее имён), пагинация — через PARAMS.NAV_PARAMS (nPageSize ≤ 50).
// Док: apidocs.bitrix24.com/api-reference/tasks/elapsed-item/task-elapsed-item-get-list.html
function tpFetchElapsed_(userId, fromApi, toApi) {
  var items = [];
  var page = 1, guard = 0;
  while (guard++ < 500) {
    var data = tpCallBitrix_('task.elapseditem.getlist', {
      ORDER: { ID: 'desc' },                       // 1 — сортировка
      FILTER: {                                     // 2 — фильтр: мои записи за период
        '>=CREATED_DATE': fromApi,
        '<=CREATED_DATE': toApi,
        'USER_ID': userId
      },
      SELECT: ['ID', 'TASK_ID', 'USER_ID', 'SECONDS', 'MINUTES',
               'COMMENT_TEXT', 'CREATED_DATE', 'DATE_START'],       // 3 — поля
      PARAMS: { NAV_PARAMS: { nPageSize: TP_PAGE_SIZE, iNumPage: page } } // 4 — постранично
    });
    var batch = data.result || [];
    if (!Array.isArray(batch)) batch = batch.items || batch.tasks || []; // на случай иной формы ответа
    items = items.concat(batch);
    if (items.length >= TP_MAX_ELAPSED) break;
    if (batch.length < TP_PAGE_SIZE) break; // последняя (неполная) страница
    page++;
  }
  return items;
}

// Данные задачи по ID (с кэшем): название, группа, исполнитель
function tpTaskInfo_(taskId, cache) {
  taskId = String(taskId);
  if (cache[taskId]) return cache[taskId];
  var info = { title: taskId, groupId: '', responsibleId: '' };
  try {
    var d = tpCallBitrix_('tasks.task.get', {
      taskId: taskId,
      select: ['ID', 'TITLE', 'GROUP_ID', 'RESPONSIBLE_ID']
    });
    var t = d.result && d.result.task;
    if (t) {
      info.title         = t.title || taskId;
      info.groupId       = t.groupId || '';
      info.responsibleId = (t.responsibleId != null) ? String(t.responsibleId) : '';
    }
  } catch (e) { /* нет доступа — оставим ID вместо названия */ }
  cache[taskId] = info;
  return info;
}

// ───────────────────── Вебхук / REST ─────────────────────
function tpUi_() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } }

function tpGetWebhook_() {
  var url = PropertiesService.getScriptProperties().getProperty(TP_PROP_WEBHOOK);
  if (!url) throw new Error('Вебхук не задан. В основном скрипте: меню «Bitrix» → «Указать вебхук».');
  return url;
}

// ID владельца вебхука = число после /rest/ в URL
function tpWebhookOwnerId_() {
  var m = tpGetWebhook_().match(/\/rest\/(\d+)\//);
  return m ? m[1] : '';
}

function tpCallBitrix_(method, params) {
  var url = tpGetWebhook_() + method + '.json';
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(params || {}),
    muteHttpExceptions: true
  });
  var code = resp.getResponseCode();
  var body = resp.getContentText();
  var data;
  try { data = JSON.parse(body); } catch (e) {
    throw new Error('Bitrix вернул не JSON (HTTP ' + code + ') для ' + method + ': ' + body.slice(0, 300));
  }
  if (data.error) {
    throw new Error('Bitrix error (' + method + '): ' + data.error + ' — ' + (data.error_description || ''));
  }
  if (code >= 400) {
    throw new Error('Bitrix HTTP ' + code + ' (' + method + '): ' + body.slice(0, 300));
  }
  return data;
}

// ───────────────────── Справочники / форматирование ─────────────────────
function tpGroupName_(id, cache) {
  if (!id || String(id) === '0') return '';
  id = String(id);
  if (cache[id]) return cache[id];
  try {
    var data = tpCallBitrix_('sonet_group.get', { FILTER: { ID: id } });
    var g = data.result && data.result[0];
    cache[id] = (g && g.NAME) ? g.NAME : id;
  } catch (e) { cache[id] = id; }
  return cache[id];
}

function tpTaskUrl_(portalBase, responsibleId, taskId) {
  if (!taskId) return '';
  return portalBase + '/company/personal/user/' + responsibleId + '/tasks/task/view/' + taskId + '/';
}

// Длительность записи в секундах: SECONDS, иначе MINUTES*60 (не суммируем — это одно и то же)
function tpSeconds_(it) {
  var sec = Number(tpPick_(it, ['SECONDS', 'seconds']));
  var mn  = Number(tpPick_(it, ['MINUTES', 'minutes']));
  if (!isNaN(sec) && sec > 0) return sec;
  if (!isNaN(mn)  && mn  > 0) return mn * 60;
  return 0;
}

// Секунды → «Ч:ММ:СС»
function tpDuration_(sec) {
  var n = Math.round(Number(sec) || 0);
  if (n < 0) n = 0;
  var h = Math.floor(n / 3600);
  var m = Math.floor((n % 3600) / 60);
  var s = n % 60;
  return h + ':' + tpPad2_(m) + ':' + tpPad2_(s);
}

function tpPad2_(x) { return (x < 10 ? '0' : '') + x; }
function tpRound1_(x) { return Math.round(x * 10) / 10; }

// Комментарий: убираем HTML-теги и лишние пробелы
function tpText_(s) {
  if (!s) return '';
  return String(s).replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '')
                  .replace(/\s+/g, ' ').trim();
}

// Первое непустое значение по списку возможных ключей (Bitrix отдаёт то ВЕРХНИЙ_РЕГИСТР, то camelCase)
function tpPick_(obj, keys) {
  for (var i = 0; i < keys.length; i++) {
    var v = obj[keys[i]];
    if (v !== undefined && v !== null) return v;
  }
  return '';
}

function tpFmtDate_(s, tz) {
  if (!s) return '';
  var d = new Date(s);
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, tz || Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

// Выравнивает длину всех строк под первую (на случай строки-заглушки)
function tpPadRows_(rows) {
  var w = rows[0].length;
  return rows.map(function (r) { while (r.length < w) r.push(''); return r; });
}
