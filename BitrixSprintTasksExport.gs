/**
 * Выгрузка всех задач ОДНОГО спринта Bitrix24 (Scrum) на лист Google Sheets.
 *
 * По умолчанию ищет спринт с названием «Спринт июль 2» (переменная SE_SPRINT_NAME
 * ниже) во всех Scrum-группах, до которых достаёт вебхук, и выгружает ВСЕ его
 * задачи (не только ваши) с планом/фактом времени, стори-поинтами и ссылкой.
 *
 * Отличие от BitrixTasksToSheets.gs: тот тянет ВАШИ задачи по всем спринтам;
 * этот — ВСЕ задачи ОДНОГО названного спринта.
 *
 * Вторая кнопка меню — «Выгрузить мои отложенные задачи»: на лист «Отложенные»
 * выгружает все задачи, где ВЫ исполнитель и статус = «Отложена» (функция
 * seExportDeferredTasks).
 *
 * КАК ПОДКЛЮЧИТЬ (один раз):
 *   1. Откройте таблицу → Расширения → Apps Script. Вставьте этот файл целиком.
 *   2. Сохраните, обновите вкладку таблицы — появится меню «Спринт».
 *   3. Вебхук. Если вы уже пользовались BitrixTasksToSheets.gs /
 *      BitrixCreateSprints.gs в ЭТОМ ЖЕ проекте Apps Script — вебхук уже
 *      сохранён (общий ключ свойств BITRIX_WEBHOOK), шаг можно пропустить.
 *      Иначе: меню «Спринт» → «Указать вебхук» и вставьте входящий вебхук
 *      портала (Битрикс: Разработчикам → Другое → Входящий вебхук, право «task»),
 *      URL вида https://ВАШ-ПОРТАЛ.bitrix24.ru/rest/123/xxxxxxxxxx/
 *   4. Меню «Спринт» → «Выгрузить задачи спринта».
 *
 * ЗАПУСК ИЗ РЕДАКТОРА (без меню): выберите функцию seExportSprintTasks и
 * нажмите «Выполнить». Результат и все сообщения уйдут в журнал (Просмотр →
 * Журналы). Вебхук в этом случае должен быть уже сохранён (см. п. 3).
 *
 * НАСТРОЙКА, какой спринт тянуть:
 *   - SE_SPRINT_NAME — имя спринта ровно как в Битриксе (регистр/пробелы не важны).
 *   - SE_GROUP_ID    — можно жёстко задать ID Scrum-группы, чтобы не искать по всем
 *                      (ускоряет и снимает неоднозначность, если одноимённые
 *                      спринты есть в разных группах). Пусто — ищем автоматически.
 */

// ───────────────────── Настройки ─────────────────────
var SE_SPRINT_NAME = 'Спринт июль 2';   // какой спринт выгружаем (имя как в Битриксе)
var SE_GROUP_ID    = '';                // '' — искать во всех группах; иначе ID Scrum-группы строкой

var SE_DEST_SHEET  = 'Задачи спринта';  // лист для результата (создаётся сам, перезаписывается)
var SE_PAGE_SIZE   = 50;                // Bitrix отдаёт задачи страницами по 50
var SE_MAX_TASKS   = 2000;              // предохранитель от бесконечного цикла/лимита 6 мин

// Кнопка «Выгрузить мои отложенные задачи»
var SE_DEFERRED_SHEET  = 'Отложенные';  // лист для отложенных задач (перезаписывается)
var SE_DEFERRED_STATUS = 6;             // статус Bitrix «Отложена» = 6

// Ключи в свойствах скрипта — те же, что в соседних скриптах (переиспользуются).
var SE_PROP_WEBHOOK  = 'BITRIX_WEBHOOK';
var SE_PROP_GROUP_ID = 'BITRIX_SCRUM_GROUP_ID';

// Поля задачи, которые забираем из Bitrix
var SE_TASK_FIELDS = [
  'ID', 'TITLE', 'STATUS', 'PRIORITY',
  'CREATED_BY', 'RESPONSIBLE_ID',
  'CREATED_DATE', 'DEADLINE', 'CLOSED_DATE',
  'TIME_ESTIMATE', 'TIME_SPENT_IN_LOGS',
  'GROUP_ID'
];

// Заголовки колонок листа «Задачи спринта» (порядок = порядок столбцов)
var SE_HEADERS = [
  'ID', 'Название', 'Статус', 'Приоритет',
  'Постановщик', 'Исполнитель',
  'Создана', 'Крайний срок', 'Завершена',
  'План, ч', 'Факт, ч', 'Откл., ч', 'В срок?',
  'SP', 'Проект/Группа', 'Ссылка'
];

// Заголовки листа «Отложенные» (без стори-поинтов — они тут не нужны)
var SE_DEFERRED_HEADERS = [
  'ID', 'Название', 'Статус', 'Приоритет',
  'Постановщик', 'Исполнитель',
  'Создана', 'Крайний срок', 'Завершена',
  'План, ч', 'Факт, ч', 'Откл., ч', 'В срок?',
  'Проект/Группа', 'Ссылка'
];

// ───────────────────── Меню ─────────────────────
function onOpen() {
  var ui;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { return; }
  ui.createMenu('Спринт')
    .addItem('Выгрузить задачи спринта', 'seExportSprintTasks')
    .addItem('Выгрузить мои отложенные задачи', 'seExportDeferredTasks')
    .addSeparator()
    .addItem('Указать вебхук', 'seSetWebhook')
    .addToUi();
}

function seUi_() {
  try { return SpreadsheetApp.getUi(); } catch (e) { return null; }
}

// ───────────────────── Вебхук ─────────────────────
function seSetWebhook() {
  var ui = seUi_();
  if (!ui) throw new Error('Диалог недоступен при запуске из редактора. Вебхук уже должен быть сохранён соседним скриптом (ключ ' + SE_PROP_WEBHOOK + ').');
  var res = ui.prompt(
    'Вебхук Bitrix24',
    'Вставьте URL входящего вебхука (право «task»):\n' +
    'https://ВАШ-ПОРТАЛ.bitrix24.ru/rest/123/xxxxxxxxxx/',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;
  var url = (res.getResponseText() || '').trim();
  if (!/^https:\/\/.+\/rest\/\d+\/.+/.test(url)) {
    ui.alert('Не похоже на вебхук. Ожидается URL вида https://портал.bitrix24.ru/rest/123/xxxxxxxxxx/');
    return;
  }
  if (url.charAt(url.length - 1) !== '/') url += '/';
  PropertiesService.getScriptProperties().setProperty(SE_PROP_WEBHOOK, url);
  ui.alert('Вебхук сохранён. Теперь меню «Спринт» → «Выгрузить задачи спринта».');
}

function seGetWebhook_() {
  var url = PropertiesService.getScriptProperties().getProperty(SE_PROP_WEBHOOK);
  if (!url) throw new Error('Вебхук не задан. Меню «Спринт» → «Указать вебхук» (или сохраните его в соседнем скрипте — ключ ' + SE_PROP_WEBHOOK + ').');
  return url;
}

// ID владельца вебхука = число после /rest/ в URL
function seWebhookOwnerId_() {
  var m = seGetWebhook_().match(/\/rest\/(\d+)\//);
  return m ? m[1] : '';
}

// ───────────────────── Вызов REST Bitrix ─────────────────────
function seCallBitrix_(method, params) {
  var url = seGetWebhook_() + method + '.json';
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

// ───────────────────── Главное действие ─────────────────────
function seExportSprintTasks() {
  var ui = seUi_();
  var wanted = String(SE_SPRINT_NAME || '').trim();
  if (!wanted) throw new Error('Не задано имя спринта — заполни SE_SPRINT_NAME вверху файла.');

  // 1. Находим спринт по имени → его id и группу
  var found = seFindSprint_(wanted);
  var sprint = found.sprint;
  var groupId = String(found.groupId);
  var sprintId = String(sprint.id);

  // 2. Тянем все задачи группы (постранично) — потом отфильтруем по спринту
  var groupTasks = seFetchGroupTasks_(groupId);

  // 3. По каждой задаче узнаём спринт (entityId) и стори-поинты; оставляем нужные
  var portalBase = seGetWebhook_().replace(/\/rest\/.*/, '');
  var userCache = {}, groupCache = {};
  var rows = [SE_HEADERS];
  var kept = 0, totalSp = 0, done = 0;

  for (var i = 0; i < groupTasks.length; i++) {
    var t = groupTasks[i];
    var scrum;
    try {
      var sd = seCallBitrix_('tasks.api.scrum.task.get', { id: t.id });
      scrum = sd.result || {};
    } catch (e) { scrum = {}; }
    if (String(scrum.entityId) !== sprintId) continue; // задача не из этого спринта

    kept++;
    var sp = Number(scrum.storyPoints) || 0;
    totalSp += sp;
    if (String(t.status) === '5') done++;

    var planH = seHours_(t.timeEstimate);
    var factH = seHours_(t.timeSpentInLogs);
    var deltaH = (planH === '' || factH === '') ? '' : seRound1_(factH - planH);

    rows.push([
      t.id,
      t.title,
      seStatusName_(t.status),
      sePriorityName_(t.priority),
      seUserName_(t.createdBy, userCache),
      seUserName_(t.responsibleId, userCache),
      seFmtDate_(t.createdDate),
      seFmtDate_(t.deadline),
      seFmtDate_(t.closedDate),
      planH,
      factH,
      deltaH,
      seOnTime_(t),
      (scrum.storyPoints || ''),
      seGroupName_(t.groupId, groupCache),
      seTaskUrl_(portalBase, t.responsibleId, t.id)
    ]);
  }

  // 4. Записываем лист
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SE_DEST_SHEET) || ss.insertSheet(SE_DEST_SHEET);
  sheet.clear();

  // Шапка со сведениями о спринте
  var title = '«' + sprint.name + '» · ' + seGroupName_(groupId, groupCache) +
    ' · ' + seFmtDay_(sprint.dateStart) + '—' + seFmtDay_(sprint.dateEnd) +
    ' · статус: ' + seSprintStatus_(sprint.status);
  var summary = 'Задач: ' + kept + ' (завершено ' + done + '), SP всего: ' + seRound1_(totalSp);
  sheet.getRange(1, 1, 1, SE_HEADERS.length).merge().setValue(title)
    .setFontWeight('bold').setBackground('#e8eaf6');
  sheet.getRange(2, 1, 1, SE_HEADERS.length).merge().setValue(summary).setFontStyle('italic');

  if (rows.length === 1) rows.push(['(в спринте нет задач или они недоступны вебхуку)']);

  var startRow = 3;
  sheet.getRange(startRow, 1, rows.length, SE_HEADERS.length).setValues(sePadRows_(rows));
  sheet.getRange(startRow, 1, 1, SE_HEADERS.length).setFontWeight('bold').setBackground('#f1f8e9');
  sheet.setFrozenRows(startRow);
  sheet.autoResizeColumns(1, SE_HEADERS.length);

  var msg = 'Спринт «' + sprint.name + '»: выгружено задач ' + kept + ' (из ' +
    groupTasks.length + ' в группе). Лист «' + SE_DEST_SHEET + '».';
  Logger.log(title + '\n' + summary + '\n' + msg);
  if (ui) SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Спринт', 8);
}

// ───────────────────── Поиск спринта по имени ─────────────────────
function seFindSprint_(wantedName) {
  var key = seNorm_(wantedName);
  var candidates = seCandidateGroups_();
  var seenNames = [];

  for (var i = 0; i < candidates.length; i++) {
    var list = seSprintsOfGroup_(candidates[i]);
    for (var j = 0; j < list.length; j++) {
      var s = list[j];
      seenNames.push(s.name + ' (группа ' + candidates[i] + ')');
      if (seNorm_(s.name) === key) {
        return { sprint: s, groupId: s.groupId || candidates[i] };
      }
    }
  }

  var hint = seenNames.length
    ? '\nНайденные спринты: ' + seenNames.join('; ') + '.'
    : '\nНи одного спринта не нашлось — проверь, что вебхук видит Scrum-группу.';
  throw new Error('Спринт «' + wantedName + '» не найден.' + hint +
    '\nПодсказка: впиши точное имя в SE_SPRINT_NAME или ID группы в SE_GROUP_ID вверху файла.');
}

// Список групп-кандидатов, где искать спринт.
function seCandidateGroups_() {
  var explicit = String(SE_GROUP_ID || '').trim() ||
    (PropertiesService.getScriptProperties().getProperty(SE_PROP_GROUP_ID) || '').trim();
  if (explicit) return [explicit];

  // Иначе — группы, где у владельца вебхука есть задачи (быстрый способ найти скрам-группы).
  var owner = seWebhookOwnerId_();
  var filter = owner ? { RESPONSIBLE_ID: owner } : {};
  var set = {};
  var start = 0, guard = 0;
  while (guard++ < 200) {
    var data = seCallBitrix_('tasks.task.list', {
      filter: filter,
      select: ['ID', 'GROUP_ID'],
      order: { ID: 'desc' },
      start: start
    });
    var batch = (data.result && data.result.tasks) || [];
    for (var i = 0; i < batch.length; i++) {
      var g = batch[i].groupId;
      if (g && String(g) !== '0') set[String(g)] = true;
    }
    if (typeof data.next === 'undefined' || batch.length === 0) break;
    start = data.next;
  }
  var groups = Object.keys(set);
  if (!groups.length) {
    throw new Error('Не удалось определить Scrum-группу автоматически. Впиши ID группы в SE_GROUP_ID вверху файла.');
  }
  return groups;
}

function seSprintsOfGroup_(groupId) {
  try {
    var data = seCallBitrix_('tasks.api.scrum.sprint.list', {
      filter: { GROUP_ID: groupId },
      select: ['*']
    });
    var list = data.result;
    if (list && !Array.isArray(list)) list = list.sprints || [];
    return list || [];
  } catch (e) {
    return []; // группа не Scrum либо нет доступа
  }
}

// ───────────────────── Мои отложенные задачи ─────────────────────
function seExportDeferredTasks() {
  var ui = seUi_();
  var responsibleId = seResponsibleId_();

  // Все задачи, где я исполнитель и статус = «Отложена»
  var tasks = seFetchTasks_({
    RESPONSIBLE_ID: responsibleId,
    STATUS: SE_DEFERRED_STATUS
  });

  var portalBase = seGetWebhook_().replace(/\/rest\/.*/, '');
  var userCache = {}, groupCache = {};
  var rows = [SE_DEFERRED_HEADERS];

  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    var planH = seHours_(t.timeEstimate);
    var factH = seHours_(t.timeSpentInLogs);
    var deltaH = (planH === '' || factH === '') ? '' : seRound1_(factH - planH);
    rows.push([
      t.id,
      t.title,
      seStatusName_(t.status),
      sePriorityName_(t.priority),
      seUserName_(t.createdBy, userCache),
      seUserName_(t.responsibleId, userCache),
      seFmtDate_(t.createdDate),
      seFmtDate_(t.deadline),
      seFmtDate_(t.closedDate),
      planH,
      factH,
      deltaH,
      seOnTime_(t),
      seGroupName_(t.groupId, groupCache),
      seTaskUrl_(portalBase, t.responsibleId, t.id)
    ]);
  }

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SE_DEFERRED_SHEET) || ss.insertSheet(SE_DEFERRED_SHEET);
  sheet.clear();

  var title = 'Мои отложенные задачи (исполнитель ID ' + responsibleId +
    ', статус «Отложена») · всего: ' + tasks.length;
  sheet.getRange(1, 1, 1, SE_DEFERRED_HEADERS.length).merge().setValue(title)
    .setFontWeight('bold').setBackground('#e8eaf6');

  if (rows.length === 1) rows.push(['(отложенных задач нет)']);

  var startRow = 2;
  sheet.getRange(startRow, 1, rows.length, SE_DEFERRED_HEADERS.length).setValues(sePadRows_(rows));
  sheet.getRange(startRow, 1, 1, SE_DEFERRED_HEADERS.length).setFontWeight('bold').setBackground('#f1f8e9');
  sheet.setFrozenRows(startRow);
  sheet.autoResizeColumns(1, SE_DEFERRED_HEADERS.length);

  var msg = 'Отложенных задач выгружено: ' + tasks.length + '. Лист «' + SE_DEFERRED_SHEET + '».';
  Logger.log(title + '\n' + msg);
  if (ui) SpreadsheetApp.getActiveSpreadsheet().toast(msg, 'Спринт', 8);
}

// Исполнитель по умолчанию — владелец вебхука (ID из URL); запасной путь — user.current.
function seResponsibleId_() {
  var id = seWebhookOwnerId_();
  if (id) return id;
  var me = seCallBitrix_('user.current', {});
  return String(me.result.ID);
}

// ───────────────────── Задачи (постранично) ─────────────────────
function seFetchGroupTasks_(groupId) {
  return seFetchTasks_({ GROUP_ID: groupId });
}

// Постранично тянет tasks.task.list по произвольному фильтру.
function seFetchTasks_(filter) {
  var tasks = [];
  var start = 0;
  while (true) {
    var data = seCallBitrix_('tasks.task.list', {
      filter: filter,
      select: SE_TASK_FIELDS,
      order: { CREATED_DATE: 'desc' },
      start: start
    });
    var batch = (data.result && data.result.tasks) || [];
    tasks = tasks.concat(batch);
    if (tasks.length >= SE_MAX_TASKS) break;
    if (typeof data.next === 'undefined' || batch.length === 0) break;
    start = data.next;
  }
  return tasks;
}

// ───────────────────── Вспомогательное для листа ─────────────────────
function sePadRows_(rows) {
  var w = rows[0].length;
  return rows.map(function (r) {
    while (r.length < w) r.push('');
    return r;
  });
}

// ───────────────────── Справочники / форматирование ─────────────────────
function seNorm_(s) {
  return String(s || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function seStatusName_(s) {
  var map = { '1': 'Новая', '2': 'Ждёт выполнения', '3': 'Выполняется',
    '4': 'Ждёт контроля', '5': 'Завершена', '6': 'Отложена', '7': 'Отклонена' };
  return map[String(s)] || String(s);
}

function sePriorityName_(p) {
  var map = { '0': 'Низкий', '1': 'Обычный', '2': 'Высокий' };
  return map[String(p)] || String(p);
}

function seSprintStatus_(s) {
  var map = { 'planned': 'запланирован', 'active': 'активный', 'completed': 'завершён' };
  return map[String(s)] || String(s || '');
}

function seUserName_(id, cache) {
  if (!id) return '';
  id = String(id);
  if (cache[id]) return cache[id];
  try {
    var data = seCallBitrix_('user.get', { ID: id });
    var u = data.result && data.result[0];
    cache[id] = (u ? [u.NAME, u.LAST_NAME].filter(String).join(' ').trim() : id) || id;
  } catch (e) { cache[id] = id; }
  return cache[id];
}

function seGroupName_(id, cache) {
  if (!id || String(id) === '0') return '';
  id = String(id);
  if (cache[id]) return cache[id];
  try {
    var data = seCallBitrix_('sonet_group.get', { FILTER: { ID: id } });
    var g = data.result && data.result[0];
    cache[id] = (g && g.NAME) ? g.NAME : id;
  } catch (e) { cache[id] = id; }
  return cache[id];
}

function seTaskUrl_(portalBase, responsibleId, taskId) {
  if (!taskId) return '';
  return portalBase + '/company/personal/user/' + responsibleId + '/tasks/task/view/' + taskId + '/';
}

// Секунды → часы (округление до 0.1); пусто, если нет данных
function seHours_(sec) {
  var n = Number(sec);
  if (!sec || isNaN(n) || n === 0) return '';
  return seRound1_(n / 3600);
}

function seRound1_(x) { return Math.round(x * 10) / 10; }

// «В срок?» — как в BitrixTasksToSheets.gs
function seOnTime_(t) {
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

function seFmtDate_(s) {
  if (!s) return '';
  var d = new Date(s);
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function seFmtDay_(s) {
  if (!s) return '';
  var d = new Date(s);
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM');
}
