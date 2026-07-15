/**
 * ОТДЕЛЬНЫЙ файл для основного проекта (тот, где лист «Задачи Bitrix» и
 * меню «Bitrix» из BitrixTasksToSheets.gs).
 *
 * Кнопка «Все мои задачи (без спринтов)» выгружает на отдельный лист ВСЕ
 * задачи, где вы — исполнитель (RESPONSIBLE_ID), вне зависимости от спринтов:
 * из скрам-групп, из обычных проектов и вообще без проекта. К спринтам скрипт
 * не обращается совсем (никаких tasks.api.scrum.*) — поэтому работает быстро
 * и не зависит от того, настроен ли скрам на портале.
 *
 * По каждой задаче: статус, приоритет, постановщик, даты (создана / крайний
 * срок / завершена), план и факт по времени, признак «в срок?», проект и
 * прямая ссылка.
 *
 * Файл самостоятельный: все функции с префиксом mt, чтобы не конфликтовать
 * с основным скриптом в том же проекте Apps Script. onOpen тут НЕТ — пункт
 * меню добавляется в основной скрипт (BitrixTasksToSheets.gs), одной строкой:
 *     .addItem('Все мои задачи (без спринтов)', 'mtExportMyTasks')
 *
 * ВЕБХУК — общий с основным скриптом (свойство BITRIX_WEBHOOK). Если вебхук
 * уже сохранён через меню «Bitrix» → «Указать вебхук», повторно вводить не нужно.
 *
 * ЗАПУСК ИЗ РЕДАКТОРА (без меню): выберите функцию mtExportMyTasks и нажмите
 * «Выполнить». Результат и сообщения — в журнале (Просмотр → Журналы).
 */

// ───────────────────── Настройки ─────────────────────
var MT_DEST_SHEET     = 'Мои задачи';     // лист результата (перезаписывается)
var MT_PROP_WEBHOOK   = 'BITRIX_WEBHOOK'; // тот же ключ, что в основном скрипте
var MT_RESPONSIBLE_ID = '';               // '' — владелец вебхука; иначе Bitrix ID строкой
var MT_INCLUDE_CLOSED = true;             // false — только активные (без «Завершена»)
var MT_MAX_TASKS      = 5000;             // предохранитель от лимита 6 минут

// Поля задачи, которые забираем из Bitrix
var MT_TASK_FIELDS = [
  'ID', 'TITLE', 'STATUS', 'PRIORITY',
  'CREATED_BY', 'RESPONSIBLE_ID',
  'CREATED_DATE', 'DEADLINE', 'CLOSED_DATE',
  'TIME_ESTIMATE', 'TIME_SPENT_IN_LOGS',
  'GROUP_ID'
];

// Заголовки колонок листа (порядок = порядок столбцов)
var MT_HEADERS = [
  'ID', 'Название', 'Статус', 'Приоритет',
  'Постановщик', 'Исполнитель',
  'Создана', 'Крайний срок', 'Завершена',
  'План, ч', 'Факт, ч', 'Откл., ч', 'В срок?',
  'Проект/Группа', 'Ссылка'
];

// ───────────────────── Главное действие ─────────────────────
function mtExportMyTasks() {
  var ui = mtUi_();

  // 1. «Я» = MT_RESPONSIBLE_ID, иначе владелец вебхука, иначе user.current
  var me = String(MT_RESPONSIBLE_ID || mtWebhookOwnerId_() ||
                  mtCallBitrix_('user.current', {}).result.ID);

  // 2. Фильтр: все задачи, где я исполнитель (спринты не учитываются вообще)
  var filter = { RESPONSIBLE_ID: me };
  if (!MT_INCLUDE_CLOSED) filter['!STATUS'] = 5; // 5 = Завершена

  // 3. Постранично тянем все задачи
  var tasks = mtFetchTasks_(filter);

  // 4. Формируем строки
  var portalBase = mtGetWebhook_().replace(/\/rest\/.*/, '');
  var userCache = {}, groupCache = {};
  var rows = [MT_HEADERS];
  var open = 0, done = 0, overdue = 0;

  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    var planH = mtHours_(t.timeEstimate);
    var factH = mtHours_(t.timeSpentInLogs);
    var deltaH = (planH === '' || factH === '') ? '' : mtRound1_(factH - planH);
    var ot = mtOnTime_(t);

    if (String(t.status) === '5') done++; else open++;
    if (ot === 'просрочена' || ot === 'просрочена (не завершена)') overdue++;

    rows.push([
      t.id,
      t.title,
      mtStatusName_(t.status),
      mtPriorityName_(t.priority),
      mtUserName_(t.createdBy, userCache),
      mtUserName_(t.responsibleId, userCache),
      mtFmtDate_(t.createdDate),
      mtFmtDate_(t.deadline),
      mtFmtDate_(t.closedDate),
      planH,
      factH,
      deltaH,
      ot,
      mtGroupName_(t.groupId, groupCache),
      mtTaskUrl_(portalBase, t.responsibleId, t.id)
    ]);
  }

  // 5. Пишем лист
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(MT_DEST_SHEET) || ss.insertSheet(MT_DEST_SHEET);
  sheet.clear();

  var title = 'Все мои задачи (исполнитель ID ' + me + ')' +
    (MT_INCLUDE_CLOSED ? '' : ' · только активные') +
    ' · всего: ' + tasks.length +
    ' · в работе: ' + open +
    ' · завершено: ' + done +
    ' · просрочено: ' + overdue;
  sheet.getRange(1, 1, 1, MT_HEADERS.length).merge().setValue(title)
    .setFontWeight('bold').setBackground('#e8eaf6');

  if (rows.length === 1) rows.push(['(задач не найдено)']);

  var startRow = 2;
  sheet.getRange(startRow, 1, rows.length, MT_HEADERS.length).setValues(mtPadRows_(rows));
  sheet.getRange(startRow, 1, 1, MT_HEADERS.length).setFontWeight('bold').setBackground('#f1f8e9');
  sheet.setFrozenRows(startRow);
  sheet.autoResizeColumns(1, MT_HEADERS.length);

  var msg = 'Моих задач выгружено: ' + tasks.length + '. Лист «' + MT_DEST_SHEET + '».';
  Logger.log(title + '\n' + msg);
  if (ui) ss.toast(msg, 'Bitrix', 8);
}

// ───────────────────── Задачи (постранично) ─────────────────────
function mtFetchTasks_(filter) {
  var tasks = [];
  var start = 0, guard = 0;
  while (guard++ < 200) {
    var data = mtCallBitrix_('tasks.task.list', {
      filter: filter,
      select: MT_TASK_FIELDS,
      order: { CREATED_DATE: 'desc' },
      start: start
    });
    var batch = (data.result && data.result.tasks) || [];
    tasks = tasks.concat(batch);
    if (tasks.length >= MT_MAX_TASKS) break;
    if (typeof data.next === 'undefined' || batch.length === 0) break;
    start = data.next;
  }
  return tasks;
}

// ───────────────────── Вебхук / REST ─────────────────────
function mtUi_() { try { return SpreadsheetApp.getUi(); } catch (e) { return null; } }

function mtGetWebhook_() {
  var url = PropertiesService.getScriptProperties().getProperty(MT_PROP_WEBHOOK);
  if (!url) throw new Error('Вебхук не задан. В основном скрипте: меню «Bitrix» → «Указать вебхук».');
  return url;
}

// ID владельца вебхука = число после /rest/ в URL
function mtWebhookOwnerId_() {
  var m = mtGetWebhook_().match(/\/rest\/(\d+)\//);
  return m ? m[1] : '';
}

function mtCallBitrix_(method, params) {
  var url = mtGetWebhook_() + method + '.json';
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
function mtStatusName_(s) {
  var map = { '1': 'Новая', '2': 'Ждёт выполнения', '3': 'Выполняется',
    '4': 'Ждёт контроля', '5': 'Завершена', '6': 'Отложена', '7': 'Отклонена' };
  return map[String(s)] || String(s);
}

function mtPriorityName_(p) {
  var map = { '0': 'Низкий', '1': 'Обычный', '2': 'Высокий' };
  return map[String(p)] || String(p);
}

function mtUserName_(id, cache) {
  if (!id) return '';
  id = String(id);
  if (cache[id]) return cache[id];
  try {
    var data = mtCallBitrix_('user.get', { ID: id });
    var u = data.result && data.result[0];
    cache[id] = (u ? [u.NAME, u.LAST_NAME].filter(String).join(' ').trim() : id) || id;
  } catch (e) { cache[id] = id; }
  return cache[id];
}

function mtGroupName_(id, cache) {
  if (!id || String(id) === '0') return '';
  id = String(id);
  if (cache[id]) return cache[id];
  try {
    var data = mtCallBitrix_('sonet_group.get', { FILTER: { ID: id } });
    var g = data.result && data.result[0];
    cache[id] = (g && g.NAME) ? g.NAME : id;
  } catch (e) { cache[id] = id; }
  return cache[id];
}

function mtTaskUrl_(portalBase, responsibleId, taskId) {
  if (!taskId) return '';
  return portalBase + '/company/personal/user/' + responsibleId + '/tasks/task/view/' + taskId + '/';
}

// Секунды → часы (округление до 0.1); пусто, если нет данных
function mtHours_(sec) {
  var n = Number(sec);
  if (!sec || isNaN(n) || n === 0) return '';
  return mtRound1_(n / 3600);
}

function mtRound1_(x) { return Math.round(x * 10) / 10; }

// «В срок?»: для завершённых — дата закрытия против дедлайна;
// для незавершённых с прошедшим дедлайном — «просрочена».
function mtOnTime_(t) {
  var deadline = t.deadline ? new Date(t.deadline) : null;
  var closed   = t.closedDate ? new Date(t.closedDate) : null;
  var done     = String(t.status) === '5';
  if (done) {
    if (!deadline) return 'без срока';
    if (!closed)   return 'в срок';
    return closed.getTime() <= deadline.getTime() ? 'в срок' : 'просрочена';
  }
  if (deadline && deadline.getTime() < Date.now()) return 'просрочена (не завершена)';
  return 'в работе';
}

function mtFmtDate_(s) {
  if (!s) return '';
  var d = new Date(s);
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

// Выравнивает длину всех строк под первую (на случай строки-заглушки)
function mtPadRows_(rows) {
  var w = rows[0].length;
  return rows.map(function (r) { while (r.length < w) r.push(''); return r; });
}
