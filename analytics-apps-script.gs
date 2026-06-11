/**
 * ТОВ «Експертна Думка» — приймач аналітики калькулятора податків + щоденний дашборд.
 *
 * Налаштування після вставки коду:
 * 1) Ввести в дію → Керувати введеннями в дію → ✏️ → Нова версія → Ввести в дію
 * 2) Один раз запусти функцію edSetupDailyTrigger() (меню «Запустити»), щоб дашборд
 *    оновлювався автоматично щодня о 8:00. Дозволь доступ, коли попросить.
 * 3) Дашборд також можна оновити вручну будь-коли: запусти edBuildDashboard().
 */

// Порядок колонок у головному аркуші (Лог). Має збігатися з пінгом калькулятора.
var ED_HEADERS = [
  'Час', 'Подія', 'Мітка', 'Сума',
  'ID пристрою', 'Тип', 'Бот', 'ОС', 'Браузер',
  'Мова', 'Часовий пояс', 'Екран', 'Вікно', 'DPR',
  'Платформа', 'Ядра', "Пам'ять (GB)", 'Сенсор',
  'Джерело (referrer)', 'utm_source', 'utm_medium', 'utm_campaign',
  'Розблоковано', 'Код доступу', 'К-сть розрах.',
  'IP', 'Країна', 'Регіон', 'Місто', 'Провайдер'
];

var ED_LOG_SHEET = 'Лог';
var ED_DASH_SHEET = 'Дашборд';

function doGet(e)  { return edHandle(e); }
function doPost(e) { return edHandle(e); }

function edLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(ED_LOG_SHEET);
  if (!sh) sh = ss.getSheets()[0];
  return sh;
}

function edHandle(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var sheet = edLogSheet();

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(ED_HEADERS);
      sheet.setFrozenRows(1);
    }

    var bot = p.bot || '';
    if (!bot) {
      var org = (p.org || '').toLowerCase();
      var ref = (p.ref || '').toLowerCase();
      var ip  = (p.ip  || '').toLowerCase();
      if (org.indexOf('facebook') !== -1 || org.indexOf('meta ') !== -1) bot = 'Meta';
      else if (org.indexOf('google') !== -1) bot = 'Google';
      else if (org.indexOf('microsoft') !== -1 || org.indexOf('bing') !== -1) bot = 'Bing';
      else if (org.indexOf('amazon') !== -1 || org.indexOf('digitalocean') !== -1 ||
               org.indexOf('ovh') !== -1 || org.indexOf('hetzner') !== -1 ||
               org.indexOf('cloudflare') !== -1) bot = 'Хостинг/ЦОД';
      else if (ip.indexOf('2a03:2880') === 0) bot = 'Meta';
    }

    var ts = p.t ? new Date(Number(p.t)) : new Date();

    var row = [
      ts, p.event || '', p.label || '', p.value || '',
      p.device || '', p.dev_type || '', bot, p.os || '', p.browser || '',
      p.lang || '', p.tz || '', p.screen || '', p.viewport || '', p.dpr || '',
      p.platform || '', p.cores || '', p.mem || '', p.touch || '',
      p.ref || '', p.utm_source || '', p.utm_medium || '', p.utm_campaign || '',
      p.unlocked || '', p.code || '', p.calc_count || '',
      p.ip || '', p.country || '', p.region || '', p.city || '', p.org || ''
    ];

    sheet.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/* ============================================================
 *  ЩОДЕННИЙ ДАШБОРД
 *  По днях: нові відвідувачі, унікальні візити, розрахунки, джерела, міста, коди.
 *  Боти (колонка «Бот» заповнена) у живу статистику НЕ входять.
 * ============================================================ */

function edBuildDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var log = edLogSheet();
  var data = log.getDataRange().getValues();
  if (data.length < 2) return;

  var H = {};
  var head = data[0];
  for (var i = 0; i < head.length; i++) H[head[i]] = i;

  var iTime = H['Час'], iEvent = H['Подія'], iDevice = H['ID пристрою'],
      iBot = H['Бот'], iCountry = H['Країна'], iCity = H['Місто'],
      iRef = H['Джерело (referrer)'], iCode = H['Код доступу'];

  var firstSeen = {};
  var days = {};

  function dayKey(d) {
    if (!(d instanceof Date)) d = new Date(d);
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  function ensureDay(k) {
    if (!days[k]) days[k] = {
      visitors: {}, newVisitors: 0, calcs: 0,
      sources: {}, countries: {}, cities: {}, bots: 0, codes: {}
    };
    return days[k];
  }

  for (var r = 1; r < data.length; r++) {
    var row = data[r];
    var t = row[iTime];
    if (!t) continue;
    var k = dayKey(t);
    var d = ensureDay(k);

    var isBot = iBot != null && row[iBot] !== '' && row[iBot] != null;
    if (isBot) { d.bots++; continue; }

    var dev = row[iDevice] || '';
    var ev  = row[iEvent] || '';

    if (dev && !firstSeen[dev]) firstSeen[dev] = k;
    if (dev) d.visitors[dev] = 1;
    if (ev === 'calculation') d.calcs++;

    var ref = (row[iRef] || '').toString();
    var src = edShortSource(ref);
    d.sources[src] = (d.sources[src] || 0) + 1;

    var country = row[iCountry] || '';
    var city = row[iCity] || '';
    if (country) d.countries[country] = (d.countries[country] || 0) + 1;
    if (city) d.cities[city] = (d.cities[city] || 0) + 1;

    var code = row[iCode] || '';
    if (code) d.codes[code] = (d.codes[code] || 0) + 1;
  }

  for (var dv in firstSeen) {
    var fk = firstSeen[dv];
    if (days[fk]) days[fk].newVisitors++;
  }

  var keys = Object.keys(days).sort().reverse();
  var out = [[
    'Дата', 'Нові', 'Унік. відвідувачі', 'Розрахунки',
    'Топ джерела', 'Топ міста', 'Коди клієнтів', 'Боти (відсіяно)'
  ]];

  for (var j = 0; j < keys.length; j++) {
    var kk = keys[j], dd = days[kk];
    out.push([
      kk, dd.newVisitors, Object.keys(dd.visitors).length, dd.calcs,
      edTopMap(dd.sources, 3), edTopMap(dd.cities, 3), edTopMap(dd.codes, 5), dd.bots
    ]);
  }

  var dash = ss.getSheetByName(ED_DASH_SHEET);
  if (!dash) dash = ss.insertSheet(ED_DASH_SHEET, 0);
  dash.clearContents();
  dash.getRange(1, 1, out.length, out[0].length).setValues(out);
  dash.setFrozenRows(1);
  dash.getRange(1, 1, 1, out[0].length).setFontWeight('bold');
  for (var c = 1; c <= out[0].length; c++) dash.autoResizeColumn(c);
}

function edShortSource(ref) {
  if (!ref || ref === '(прямий)' || ref === '') return 'Прямий/застосунок';
  ref = ref.toLowerCase();
  if (ref.indexOf('instagram') !== -1) return 'Instagram';
  if (ref.indexOf('facebook') !== -1) return 'Facebook';
  if (ref.indexOf('t.me') !== -1 || ref.indexOf('telegram') !== -1) return 'Telegram';
  if (ref.indexOf('google') !== -1) return 'Google';
  if (ref.indexOf('youtube') !== -1) return 'YouTube';
  if (ref.indexOf('olx') !== -1) return 'OLX';
  if (ref.indexOf('viber') !== -1) return 'Viber';
  if (ref.indexOf('whatsapp') !== -1) return 'WhatsApp';
  try {
    var m = ref.match(/^https?:\/\/([^\/]+)/);
    return m ? m[1] : ref;
  } catch (e) { return ref; }
}

function edTopMap(map, n) {
  var arr = [];
  for (var k in map) arr.push([k, map[k]]);
  arr.sort(function (a, b) { return b[1] - a[1]; });
  arr = arr.slice(0, n);
  return arr.map(function (x) { return x[0] + ' (' + x[1] + ')'; }).join(', ');
}

function edSetupDailyTrigger() {
  var trs = ScriptApp.getProjectTriggers();
  for (var i = 0; i < trs.length; i++) {
    if (trs[i].getHandlerFunction() === 'edBuildDashboard') ScriptApp.deleteTrigger(trs[i]);
  }
  ScriptApp.newTrigger('edBuildDashboard').timeBased().everyDays(1).atHour(8).create();
  edBuildDashboard();
}
