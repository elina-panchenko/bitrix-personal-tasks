/**
 * ОТДЕЛЬНЫЙ скрипт для ОТДЕЛЬНОЙ новой таблицы.
 * Выгружает из Bitrix24 задачи, в которых вы — исполнитель (RESPONSIBLE_ID),
 * на лист «Задачи Bitrix» текущей таблицы.
 *
 * Не связан с Kod.gs — это самостоятельный файл для нового документа.
 *
 * КАК ПОДКЛЮЧИТЬ (один раз):
 *   1. Создайте НОВУЮ Google-таблицу.
 *   2. Расширения → Apps Script. Вставьте этот файл целиком.
 *   3. Сохраните. Обновите вкладку с таблицей — появится меню «Bitrix».
 *   4. Меню «Bitrix» → «Указать вебхук» — вставьте входящий вебхук портала
 *      (Битрикс: Разработчикам → Другое → Входящий вебхук, право «task»),
 *      URL вида: https://ВАШ-ПОРТАЛ.bitrix24.ru/rest/123/xxxxxxxxxx/
 *      Вебхук хранится в свойствах скрипта, в коде и в таблице его нет.
 *   5. Меню «Bitrix» → «Обновить задачи». Первый раз Google попросит
 *      разрешения (скрипт работает от вашего аккаунта) — разрешите.
 *
 * Необязательно: меню → «Включить автообновление» поставит ежедневный
 * триггер (каждое утро лист будет обновляться сам).
 */

// ───────────────────── Настройки ─────────────────────
var DEST_SHEET   = 'Задачи Bitrix';   // лист, куда пишем результат (создаётся сам)
var PAGE_SIZE    = 50;                 // Bitrix отдаёт задачи страницами по 50
var PROP_WEBHOOK = 'BITRIX_WEBHOOK';   // ключ в свойствах скрипта с URL вебхука

// Чьи задачи выгружать:
//   ''            — текущий пользователь вебхука (определяется через user.current)
//   '123'         — конкретный Bitrix ID исполнителя (строкой)
var RESPONSIBLE_ID = '';

// Включать ли завершённые/закрытые задачи:
//   true  — все задачи (включая завершённые)
//   false — только активные (статус != «Завершена»)
var INCLUDE_CLOSED = true;

// Скрам: тянуть ли спринты и стори-поинты и строить ретроспективу.
var SCRUM_ENABLED = true;
// Предохранитель: максимум обращений за стори-поинтами (защита от лимита 6 мин).
var MAX_SCRUM_LOOKUPS = 400;

// Имена служебных листов
var SHEET_SPRINTS  = 'Спринты';
var SHEET_RETRO    = 'Ретро';
var SHEET_PLANNING = 'Планирование';

// Стартовые параметры ёмкости (используются при ПЕРВОМ создании листа
// «Планирование»; потом правьте прямо в таблице — скрипт их не перетирает).
var CAP_DAYS_IN_SPRINT   = 5;    // рабочих дней в спринте
var CAP_FOCUS_HOURS_DAY  = 3.5;  // фокус-часов в день под ваши задачи (3–4)
var CAP_TECHDEBT_PCT     = 20;   // резерв на техдолг, %
var CAP_BUFFER_PCT       = 15;   // резерв на непредвиденное/контекст-свитчинг, %
var CAP_ACCURACY         = 1.3;  // средняя точность оценок (факт/план), >1 = недооценка
var CAP_HOURS_PER_SP     = 4;    // сколько часов в 1 стори-поинте (ваша калибровка)

// Поля задачи, которые забираем из Bitrix
var TASK_FIELDS = [
  'ID', 'TITLE', 'STATUS', 'PRIORITY',
  'CREATED_BY', 'RESPONSIBLE_ID',
  'CREATED_DATE', 'DEADLINE', 'CLOSED_DATE',
  'TIME_ESTIMATE', 'TIME_SPENT_IN_LOGS',
  'GROUP_ID'
];

// Заголовки колонок на листе (порядок = порядок столбцов)
var HEADERS = [
  'ID', 'Название', 'Статус', 'Приоритет',
  'Постановщик', 'Исполнитель',
  'Создана', 'Крайний срок', 'Завершена',
  'План, ч', 'Факт, ч', 'Откл., ч', 'В срок?',
  'Спринт', 'SP', 'Проект/Группа', 'Ссылка'
];

// ───────────────────── Меню ─────────────────────
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Bitrix')
    .addItem('Обновить задачи + ретро', 'updateBitrixTasks')
    .addItem('Выгрузить за сегодня (время / выполнено)', 'exportTodayTimeAndDone')
    .addItem('Учёт времени за период', 'tpExportTimeByPeriod')
    .addItem('Пересоздать лист «Планирование»', 'rebuildPlanningSheet')
    .addSeparator()
    .addItem('Указать вебхук', 'setBitrixWebhook')
    .addItem('Включить автообновление (ежедневно)', 'enableDailyTrigger')
    .addItem('Выключить автообновление', 'disableDailyTrigger')
    .addToUi();
}

// ───────────────────── Вебхук ─────────────────────
function setBitrixWebhook() {
  var ui = SpreadsheetApp.getUi();
  var res = ui.prompt(
    'Вебхук Bitrix24',
    'Вставьте URL входящего вебхука (право «task»):\n' +
    'https://ВАШ-ПОРТАЛ.bitrix24.ru/rest/123/xxxxxxxxxx/',
    ui.ButtonSet.OK_CANCEL
  );
  if (res.getSelectedButton() !== ui.Button.OK) return;

  var url = (res.getResponseText() || '').trim();
  if (!/^https:\/\/.+\/rest\/\d+\/.+/.test(url)) {
    ui.alert('Не похоже на вебхук. Ожидается URL вида ' +
             'https://портал.bitrix24.ru/rest/123/xxxxxxxxxx/');
    return;
  }
  if (url.charAt(url.length - 1) !== '/') url += '/';
  PropertiesService.getScriptProperties().setProperty(PROP_WEBHOOK, url);
  ui.alert('Вебхук сохранён. Теперь меню «Bitrix» → «Обновить задачи».');
}

// ID владельца вебхука = число после /rest/ в URL (это ваш Bitrix ID)
function webhookOwnerId_() {
  var m = getWebhook_().match(/\/rest\/(\d+)\//);
  return m ? m[1] : '';
}

function getWebhook_() {
  var url = PropertiesService.getScriptProperties().getProperty(PROP_WEBHOOK);
  if (!url) {
    throw new Error('Вебхук не задан. Меню «Bitrix» → «Указать вебхук».');
  }
  return url;
}

// ───────────────────── Вызов REST Bitrix ─────────────────────
function callBitrix_(method, params) {
  var url = getWebhook_() + method + '.json';
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
    throw new Error('Bitrix вернул не JSON (HTTP ' + code + '): ' + body.slice(0, 300));
  }
  if (data.error) {
    throw new Error('Bitrix error: ' + data.error + ' — ' + (data.error_description || ''));
  }
  if (code >= 400) {
    throw new Error('Bitrix HTTP ' + code + ': ' + body.slice(0, 300));
  }
  return data;
}

// ───────────────────── Главное действие ─────────────────────
function updateBitrixTasks() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Определяем исполнителя
  //    Если RESPONSIBLE_ID не задан — берём ID владельца вебхука прямо из URL
  //    (число после /rest/). Это не требует права «user». Если не вышло —
  //    запасной вариант через user.current (нужно право «user»).
  var responsibleId = RESPONSIBLE_ID;
  if (!responsibleId) {
    responsibleId = webhookOwnerId_();
  }
  if (!responsibleId) {
    var me = callBitrix_('user.current', {});
    responsibleId = String(me.result.ID);
  }

  // 2. Фильтр
  var filter = { RESPONSIBLE_ID: responsibleId };
  if (!INCLUDE_CLOSED) {
    filter['!STATUS'] = 5; // 5 = Завершена
  }

  // 3. Постранично тянем все задачи
  var tasks = [];
  var start = 0;
  while (true) {
    var data = callBitrix_('tasks.task.list', {
      filter: filter,
      select: TASK_FIELDS,
      order: { CREATED_DATE: 'desc' },
      start: start
    });
    var batch = (data.result && data.result.tasks) || [];
    tasks = tasks.concat(batch);
    if (typeof data.next === 'undefined' || batch.length === 0) break;
    start = data.next;
    if (start > 100000) break; // страховка от бесконечного цикла
  }

  // 4. Подтягиваем имена пользователей и названий групп (с кэшем)
  var userCache  = {};
  var groupCache = {};

  // 4b. Скрам: карта спринтов и стори-поинты по задачам
  var sprintMap   = {};   // entityId -> {id,name,goal,dateStart,dateEnd,status,groupId}
  var scrumGroups = {};   // groupId -> true (группы, где есть спринты)
  var scrumByTask = {};   // taskId -> {entityId, storyPoints}
  if (SCRUM_ENABLED) {
    sprintMap   = fetchSprints_(distinctGroupIds_(tasks), scrumGroups);
    scrumByTask = fetchScrumTasks_(tasks, scrumGroups);
  }

  // 5. Формируем строки
  var portalBase = getWebhook_().replace(/\/rest\/.*/, '');
  var rows = [HEADERS];
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    var planH = hours_(t.timeEstimate);
    var factH = hours_(t.timeSpentInLogs);
    var deltaH = (planH === '' || factH === '') ? '' : round1_(factH - planH);
    var scrum = scrumByTask[String(t.id)] || {};
    rows.push([
      t.id,
      t.title,
      statusName_(t.status),
      priorityName_(t.priority),
      userName_(t.createdBy, userCache),
      userName_(t.responsibleId, userCache),
      fmtDate_(t.createdDate),
      fmtDate_(t.deadline),
      fmtDate_(t.closedDate),
      planH,
      factH,
      deltaH,
      onTime_(t),
      sprintLabel_(scrum.entityId, sprintMap),
      (scrum.storyPoints || ''),
      groupName_(t.groupId, groupCache),
      taskUrl_(portalBase, t.responsibleId, t.id)
    ]);
  }

  // 6. Перезаписываем лист задач
  var sheet = ss.getSheetByName(DEST_SHEET) || ss.insertSheet(DEST_SHEET);
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, HEADERS.length).setValues(rows);
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, HEADERS.length);

  // 7. Служебные листы
  if (SCRUM_ENABLED) {
    buildSprintsSheet_(ss, sprintMap, groupCache);
    buildRetroSheet_(ss, tasks, scrumByTask, sprintMap);
  }
  buildPlanningSheet_(ss, false); // создаст, только если листа ещё нет

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Выгружено задач: ' + tasks.length +
    (SCRUM_ENABLED ? ', спринтов: ' + Object.keys(sprintMap).length : ''),
    'Bitrix', 6);
}

// ───────────────────── Скрам ─────────────────────
function distinctGroupIds_(tasks) {
  var set = {};
  for (var i = 0; i < tasks.length; i++) {
    var g = tasks[i].groupId;
    if (g && String(g) !== '0') set[String(g)] = true;
  }
  return Object.keys(set);
}

// Спринты по всем группам задач. Заполняет scrumGroups (где спринты нашлись).
function fetchSprints_(groupIds, scrumGroups) {
  var map = {};
  for (var i = 0; i < groupIds.length; i++) {
    var gid = groupIds[i];
    try {
      var data = callBitrix_('tasks.api.scrum.sprint.list', {
        filter: { GROUP_ID: gid },
        select: ['*']
      });
      var list = data.result;
      if (list && !Array.isArray(list)) list = list.sprints || [];
      list = list || [];
      if (list.length) scrumGroups[gid] = true;
      for (var j = 0; j < list.length; j++) {
        var s = list[j];
        map[String(s.id)] = s;
      }
    } catch (e) { /* группа не скрам — пропускаем */ }
  }
  return map;
}

// Стори-поинты и принадлежность спринту — по каждой задаче скрам-групп.
function fetchScrumTasks_(tasks, scrumGroups) {
  var out = {};
  var n = 0;
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    if (!scrumGroups[String(t.groupId)]) continue;
    if (n >= MAX_SCRUM_LOOKUPS) break;
    n++;
    try {
      var data = callBitrix_('tasks.api.scrum.task.get', { id: t.id });
      var r = data.result || {};
      out[String(t.id)] = { entityId: r.entityId, storyPoints: r.storyPoints };
    } catch (e) { /* нет скрам-данных по задаче */ }
  }
  return out;
}

function sprintLabel_(entityId, sprintMap) {
  if (!entityId) return '';
  var s = sprintMap[String(entityId)];
  return s ? s.name : 'бэклог';
}

// Лист «Спринты»
function buildSprintsSheet_(ss, sprintMap, groupCache) {
  var head = ['ID', 'Спринт', 'Цель', 'Статус', 'Начало', 'Конец', 'Группа'];
  var rows = [head];
  var ids = Object.keys(sprintMap).sort(function (a, b) {
    return String(sprintMap[a].dateStart || '').localeCompare(String(sprintMap[b].dateStart || ''));
  });
  for (var i = 0; i < ids.length; i++) {
    var s = sprintMap[ids[i]];
    rows.push([
      s.id, s.name, s.goal || '', sprintStatus_(s.status),
      fmtDate_(s.dateStart), fmtDate_(s.dateEnd),
      groupName_(s.groupId, groupCache)
    ]);
  }
  writeSheet_(ss, SHEET_SPRINTS, rows, head.length);
}

function sprintStatus_(s) {
  var map = { 'planned': 'запланирован', 'active': 'активный', 'completed': 'завершён' };
  return map[String(s)] || String(s || '');
}

// Лист «Ретро»: агрегаты по каждому спринту
function buildRetroSheet_(ss, tasks, scrumByTask, sprintMap) {
  var agg = {};  // entityId -> агрегаты
  function bucket(id) {
    if (!agg[id]) agg[id] = {
      tasks: 0, done: 0, spPlan: 0, spDone: 0,
      planH: 0, factH: 0, overdue: 0
    };
    return agg[id];
  }
  for (var i = 0; i < tasks.length; i++) {
    var t = tasks[i];
    var sc = scrumByTask[String(t.id)];
    if (!sc || !sc.entityId || !sprintMap[String(sc.entityId)]) continue; // только задачи спринтов
    var b = bucket(String(sc.entityId));
    var done = String(t.status) === '5';
    var sp = Number(sc.storyPoints) || 0;
    var ph = Number(hours_(t.timeEstimate)) || 0;
    var fh = Number(hours_(t.timeSpentInLogs)) || 0;
    b.tasks++;
    if (done) { b.done++; b.spDone += sp; }
    b.spPlan += sp;
    b.planH += ph;
    b.factH += fh;
    var ot = onTime_(t);
    if (ot === 'просрочена' || ot === 'просрочена (не завершена)') b.overdue++;
  }

  var head = ['Спринт', 'Период', 'Статус', 'Задач', 'Завершено', '% задач',
              'SP план', 'SP сделано', '% SP', 'План, ч', 'Факт, ч',
              'Точность (факт/план)', 'Просрочено'];
  var rows = [head];
  var ids = Object.keys(agg).sort(function (a, b) {
    return String(sprintMap[a].dateStart || '').localeCompare(String(sprintMap[b].dateStart || ''));
  });
  for (var k = 0; k < ids.length; k++) {
    var s = sprintMap[ids[k]];
    var a = agg[ids[k]];
    rows.push([
      s.name,
      fmtDay_(s.dateStart) + ' — ' + fmtDay_(s.dateEnd),
      sprintStatus_(s.status),
      a.tasks, a.done, pct_(a.done, a.tasks),
      round1_(a.spPlan), round1_(a.spDone), pct_(a.spDone, a.spPlan),
      round1_(a.planH), round1_(a.factH),
      a.planH ? round1_(a.factH / a.planH) : '',
      a.overdue
    ]);
  }
  if (rows.length === 1) rows.push(['(нет задач, привязанных к спринтам)']);
  writeSheet_(ss, SHEET_RETRO, rows, head.length);
}

// Лист «Планирование»: калькулятор ёмкости (создаётся один раз, потом не трётся)
function rebuildPlanningSheet() {
  buildPlanningSheet_(SpreadsheetApp.getActiveSpreadsheet(), true);
  SpreadsheetApp.getUi().alert('Лист «Планирование» пересоздан со стартовыми параметрами.');
}

function buildPlanningSheet_(ss, force) {
  var existing = ss.getSheetByName(SHEET_PLANNING);
  if (existing && !force) return;            // не перетираем правки пользователя
  var sheet = existing || ss.insertSheet(SHEET_PLANNING);
  sheet.clear();

  var data = [
    ['ПЛАНИРОВАНИЕ СПРИНТА — ёмкость', ''],
    ['Рабочих дней в спринте', CAP_DAYS_IN_SPRINT],
    ['Фокус-часов в день (только мои задачи)', CAP_FOCUS_HOURS_DAY],
    ['Валовая фокус-ёмкость, ч', '=B2*B3'],
    ['Резерв на техдолг, %', CAP_TECHDEBT_PCT],
    ['Резерв на непредвиденное/контекст-свитчинг, %', CAP_BUFFER_PCT],
    ['Ёмкость под плановые задачи, ч', '=B4*(1-B5/100-B6/100)'],
    ['Средняя точность оценок (факт/план)', CAP_ACCURACY],
    ['Рекомендуемый объём план-часов к набору', '=B7/B8'],
    ['Часов в одном стори-поинте (ваша калибровка)', CAP_HOURS_PER_SP],
    ['Рекомендуемый объём SP к набору', '=B9/B10'],
    ['', ''],
    ['Подсказка', 'Правьте синие значения в столбце B — формулы пересчитаются. ' +
     'Точность оценок и часы/SP сверяйте с листом «Ретро».']
  ];
  sheet.getRange(1, 1, data.length, 2).setValues(data);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#e8eaf6');
  // выделяем редактируемые входные ячейки
  [2, 3, 5, 6, 8, 10].forEach(function (r) {
    sheet.getRange(r, 2).setBackground('#e3f2fd');
  });
  // выделяем итоговые
  [4, 7, 9, 11].forEach(function (r) {
    sheet.getRange(r, 2).setFontWeight('bold').setBackground('#f1f8e9');
  });
  sheet.autoResizeColumns(1, 2);
  sheet.setColumnWidth(1, 320);
}

// ───────────────────── Вспомогательное для листов ─────────────────────
function writeSheet_(ss, name, rows, cols) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(padRows_(rows));
  sheet.getRange(1, 1, 1, cols).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, cols);
}

// выравнивает длину строк под первую (на случай строки-заглушки)
function padRows_(rows) {
  var w = rows[0].length;
  return rows.map(function (r) {
    while (r.length < w) r.push('');
    return r;
  });
}

function pct_(part, whole) {
  if (!whole) return '';
  return Math.round((part / whole) * 100) + '%';
}

function fmtDay_(s) {
  if (!s) return '';
  var d = new Date(s);
  if (isNaN(d.getTime())) return s;
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd.MM');
}

// ───────────────────── Справочники ─────────────────────
function statusName_(s) {
  var map = {
    '1': 'Новая',
    '2': 'Ждёт выполнения',
    '3': 'Выполняется',
    '4': 'Ждёт контроля',
    '5': 'Завершена',
    '6': 'Отложена',
    '7': 'Отклонена'
  };
  return map[String(s)] || String(s);
}

function priorityName_(p) {
  var map = { '0': 'Низкий', '1': 'Обычный', '2': 'Высокий' };
  return map[String(p)] || String(p);
}

function userName_(id, cache) {
  if (!id) return '';
  id = String(id);
  if (cache[id]) return cache[id];
  try {
    var data = callBitrix_('user.get', { ID: id });
    var u = data.result && data.result[0];
    var name = u ? [u.NAME, u.LAST_NAME].filter(String).join(' ').trim() : id;
    cache[id] = name || id;
  } catch (e) {
    cache[id] = id;
  }
  return cache[id];
}

function groupName_(id, cache) {
  if (!id || String(id) === '0') return '';
  id = String(id);
  if (cache[id]) return cache[id];
  try {
    var data = callBitrix_('sonet_group.get', { FILTER: { ID: id } });
    var g = data.result && data.result[0];
    cache[id] = (g && g.NAME) ? g.NAME : id;
  } catch (e) {
    cache[id] = id;
  }
  return cache[id];
}

function taskUrl_(portalBase, responsibleId, taskId) {
  if (!taskId) return '';
  return portalBase + '/company/personal/user/' + responsibleId +
         '/tasks/task/view/' + taskId + '/';
}

// Секунды → часы (с округлением до 0.1); пусто, если нет данных
function hours_(sec) {
  var n = Number(sec);
  if (!sec || isNaN(n) || n === 0) return '';
  return round1_(n / 3600);
}

function round1_(x) {
  return Math.round(x * 10) / 10;
}

// «В срок?»: для завершённых — сравнение даты закрытия с дедлайном;
// для незавершённых с прошедшим дедлайном — «просрочена».
function onTime_(t) {
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

function fmtDate_(s) {
  if (!s) return '';
  var d = new Date(s);
  if (isNaN(d.getTime())) return s;
  var tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm');
}

// ───────────────────── Автообновление ─────────────────────
function enableDailyTrigger() {
  disableDailyTrigger();
  ScriptApp.newTrigger('updateBitrixTasks')
    .timeBased().everyDays(1).atHour(7).create();
  SpreadsheetApp.getUi().alert('Автообновление включено: каждый день ~07:00.');
}

function disableDailyTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'updateBitrixTasks') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}

// ═════════════════ Сегодня: время / выполнено ═════════════════
// Отдельное меню-действие. Пишет ДВА листа за сегодня (по отдельности):
//   • «Сегодня: учёт времени» — задачи, где вы сегодня вели учёт времени
//     (есть записи task.elapseditem, созданные сегодня вами; остановленный
//     таймер тоже создаёт такую запись);
//   • «Сегодня: выполнено» — задачи, которые вы сегодня закрыли
//     (CLOSED_DATE = сегодня, статус «Завершена»; учитываются и те, где вы —
//     тот, кто закрыл, даже если исполнитель другой).
// Вебхук, вызов REST и справочники — общие с основным скриптом (тот же
// сохранённый BITRIX_WEBHOOK), повторно вводить вебхук не нужно.

var TD_SHEET_TIME   = 'Сегодня: учёт времени';       // сводка по задачам (создаётся сам)
var TD_SHEET_DETAIL = 'Сегодня: детализация времени'; // каждая запись времени отдельной строкой
var TD_SHEET_DONE   = 'Сегодня: выполнено';          // задачи, закрытые сегодня
var TD_TIME_BASIS = 'CREATED';  // 'CREATED' — по дате записи времени;
                                // 'STARTED' — по дате старта таймера (DATE_START)
var TD_TIMEZONE   = '';         // '' — пояс скрипта; иначе напр. 'Europe/Moscow'

var TD_TASK_FIELDS = [
  'ID', 'TITLE', 'STATUS', 'RESPONSIBLE_ID', 'CREATED_BY',
  'CLOSED_BY', 'CLOSED_DATE', 'CHANGED_DATE', 'ACTIVITY_DATE',
  'GROUP_ID', 'TIME_SPENT_IN_LOGS'
];

var TD_TIME_HEADERS = [
  'ID', 'Название', 'Статус',
  'Время сегодня, мин', 'Время сегодня, ч', 'Записей времени',
  'Также завершена сегодня?', 'Проект/Группа', 'Ссылка'
];

// Детализация: одна строка = одна запись учёта времени за сегодня
var TD_DETAIL_HEADERS = [
  'Дата и время', 'Длительность', 'Длительность, мин', 'Комментарий',
  'Задача', 'ID задачи', 'Проект/Группа', 'Ссылка'
];

var TD_DONE_HEADERS = [
  'ID', 'Название', 'Статус', 'Завершена (дата)', 'Кто завершил',
  'Время сегодня, мин', 'Записей времени', 'Проект/Группа', 'Ссылка'
];

function exportTodayTimeAndDone() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = TD_TIMEZONE || Session.getScriptTimeZone();

  // «Я» — как и в updateBitrixTasks: RESPONSIBLE_ID, иначе владелец вебхука.
  var me = RESPONSIBLE_ID || webhookOwnerId_();
  if (!me) me = String(callBitrix_('user.current', {}).result.ID);
  me = String(me);

  var todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  // Кандидаты: задачи, которых сегодня что-то коснулось. Для «времени» —
  // мои задачи с сегодняшней активностью/изменением. Для «закрытых» — где
  // сегодня выставлен CLOSED_DATE и я исполнитель ИЛИ я закрыл(а).
  var candidates = {};
  tdCollect_({ RESPONSIBLE_ID: me, '>=ACTIVITY_DATE': todayStr }, candidates);
  tdCollect_({ RESPONSIBLE_ID: me, '>=CHANGED_DATE':  todayStr }, candidates);
  tdCollect_({ RESPONSIBLE_ID: me, '>=CLOSED_DATE':   todayStr }, candidates);
  tdCollect_({ CLOSED_BY: me,      '>=CLOSED_DATE':   todayStr }, candidates);

  var userCache = {}, groupCache = {};
  var portalBase = getWebhook_().replace(/\/rest\/.*/, '');
  var timeRows = [], detailRows = [], doneRows = [];
  var totalSec = 0;

  Object.keys(candidates).forEach(function (id) {
    var t = candidates[id];
    var time = tdTodayTime_(id, me, todayStr, tz);   // {minutes, count, items[]}
    var doneToday = tdDoneToday_(t, todayStr, tz);
    var groupTitle = groupName_(t.groupId, groupCache);
    var url = taskUrl_(portalBase, t.responsibleId, t.id);

    if (time.count > 0) {
      timeRows.push([
        t.id, t.title, statusName_(t.status),
        time.minutes, round1_(time.minutes / 60), time.count,
        doneToday ? 'да' : 'нет', groupTitle, url
      ]);
      // строки детализации — по каждой записи времени за сегодня
      time.items.forEach(function (it) {
        totalSec += it.seconds;
        detailRows.push([
          tdFmt_(it.when, tz),
          tdHMS_(it.seconds),
          round1_(it.seconds / 60),
          it.comment,
          t.title, t.id, groupTitle, url
        ]);
      });
    }
    if (doneToday) {
      doneRows.push([
        t.id, t.title, statusName_(t.status),
        tdFmt_(t.closedDate, tz),
        userName_(t.closedBy || t.responsibleId, userCache),
        time.minutes, time.count, groupTitle, url
      ]);
    }
  });

  timeRows.sort(function (a, b) { return b[3] - a[3]; });                       // больше минут выше
  detailRows.sort(function (a, b) { return String(b[0]).localeCompare(String(a[0])); }); // позже записанные выше
  doneRows.sort(function (a, b) { return String(b[3]).localeCompare(String(a[3])); }); // позже закрытые выше

  tdWriteSheet_(ss, TD_SHEET_TIME, TD_TIME_HEADERS, timeRows,
    '(за ' + todayStr + ' учёт времени не запускался)');
  tdWriteSheet_(ss, TD_SHEET_DETAIL, TD_DETAIL_HEADERS, detailRows,
    '(за ' + todayStr + ' записей времени нет)');
  tdWriteSheet_(ss, TD_SHEET_DONE, TD_DONE_HEADERS, doneRows,
    '(за ' + todayStr + ' завершённых задач нет)');

  ss.toast('За ' + todayStr + ': записей времени — ' + detailRows.length +
           ' (' + tdHMS_(totalSec) + '), закрытых задач — ' + doneRows.length, 'Bitrix', 8);
}

// Пишет лист: шапка + строки, либо строка-заглушка, если строк нет.
function tdWriteSheet_(ss, name, headers, dataRows, emptyMsg) {
  var w = headers.length;
  var rows = [headers].concat(dataRows);
  if (dataRows.length === 0) rows.push([emptyMsg]);
  rows = rows.map(function (r) { while (r.length < w) r.push(''); return r; });
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, w).setValues(rows);
  sheet.getRange(1, 1, 1, w).setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, w);
}

// Постранично тянет задачи по фильтру в map по ID (без дублей)
function tdCollect_(filter, into) {
  var start = 0;
  while (true) {
    var data = callBitrix_('tasks.task.list', {
      filter: filter, select: TD_TASK_FIELDS, order: { ID: 'desc' }, start: start
    });
    var batch = (data.result && data.result.tasks) || [];
    for (var i = 0; i < batch.length; i++) into[String(batch[i].id)] = batch[i];
    if (typeof data.next === 'undefined' || batch.length === 0) break;
    start = data.next;
    if (start > 100000) break;
  }
}

// Записи затраченного времени по задаче за сегодня этим пользователем.
// Глобального фильтра по дате у task.elapseditem нет — берём записи задачи
// и фильтруем на стороне скрипта.
function tdTodayTime_(taskId, me, todayStr, tz) {
  var minutes = 0, out = [], data;
  try {
    data = callBitrix_('task.elapseditem.getlist', {
      TASKID: taskId, ORDER: { ID: 'desc' }
    });
  } catch (e) { return { minutes: 0, count: 0, items: [] }; }
  var items = data.result;
  if (items && !Array.isArray(items)) items = items.items || [];
  items = items || [];
  for (var i = 0; i < items.length; i++) {
    var it = items[i];
    if (String(it.USER_ID) !== me) continue;
    var basis = (TD_TIME_BASIS === 'STARTED') ? it.DATE_START : it.CREATED_DATE;
    if (tdDay_(basis, tz) !== todayStr) continue;
    // SECONDS и MINUTES — одна и та же длительность в разных единицах,
    // не суммируем: берём секунды, иначе минуты.
    var sec = Number(it.SECONDS), mn = Number(it.MINUTES);
    var itemSec = (!isNaN(sec) && sec > 0) ? sec
                : (!isNaN(mn) && mn > 0)   ? mn * 60 : 0;
    minutes += itemSec / 60;
    out.push({
      when: basis,
      seconds: itemSec,
      comment: tdText_(it.COMMENT_TEXT)
    });
  }
  return { minutes: round1_(minutes), count: out.length, items: out };
}

// Секунды → «Ч:ММ:СС»
function tdHMS_(totalSec) {
  var s = Math.round(Number(totalSec) || 0);
  var h = Math.floor(s / 3600);
  var m = Math.floor((s % 3600) / 60);
  var ss = s % 60;
  function p(n) { return (n < 10 ? '0' : '') + n; }
  return h + ':' + p(m) + ':' + p(ss);
}

// Комментарий записи: убираем HTML-теги и лишние пробелы
function tdText_(s) {
  if (!s) return '';
  return String(s).replace(/<br\s*\/?>/gi, ' ').replace(/<[^>]+>/g, '')
                  .replace(/\s+/g, ' ').trim();
}

function tdDoneToday_(t, todayStr, tz) {
  if (String(t.status) !== '5') return false;      // 5 = Завершена
  if (!t.closedDate) return false;
  return tdDay_(t.closedDate, tz) === todayStr;
}

function tdDay_(s, tz) {
  if (!s) return '';
  var d = new Date(s);
  return isNaN(d.getTime()) ? '' : Utilities.formatDate(d, tz, 'yyyy-MM-dd');
}

function tdFmt_(s, tz) {
  if (!s) return '';
  var d = new Date(s);
  return isNaN(d.getTime()) ? s : Utilities.formatDate(d, tz, 'yyyy-MM-dd HH:mm');
}
