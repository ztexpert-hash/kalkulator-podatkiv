/**
 * ТОВ «Експертна Думка» — приймач аналітики калькулятора податків.
 * Розгортання: Apps Script → Розгорнути → Веб-застосунок →
 *   Виконувати як: Я (твій акаунт)
 *   Хто має доступ: Будь-хто
 * Після зміни коду ОБОВ'ЯЗКОВО створюй НОВЕ розгортання (або «Керувати розгортаннями» → редагувати → нова версія),
 * інакше старий URL працюватиме зі старим кодом.
 *
 * Якщо змінюєш набір колонок — видали рядок-заголовок у таблиці,
 * скрипт створить новий автоматично при першому записі.
 */

// Порядок колонок у таблиці. Часова мітка — перша.
var ED_HEADERS = [
  'Час', 'Подія', 'Мітка', 'Сума',
  'ID пристрою', 'Тип', 'ОС', 'Браузер',
  'Мова', 'Часовий пояс', 'Екран', 'Вікно', 'DPR',
  'Платформа', 'Ядра', "Пам'ять (GB)", 'Сенсор',
  'Джерело (referrer)', 'utm_source', 'utm_medium', 'utm_campaign',
  'Розблоковано', 'Код доступу', 'К-сть розрах.',
  'IP', 'Країна', 'Регіон', 'Місто', 'Провайдер'
];

function doGet(e)  { return edHandle(e); }
function doPost(e) { return edHandle(e); }

function edHandle(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

    // Гарантуємо заголовок
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(ED_HEADERS);
      sheet.setFrozenRows(1);
    }

    // IP/гео тепер приходять з клієнта (через ipapi.co), не з сервера
    var ts = p.t ? new Date(Number(p.t)) : new Date();

    var row = [
      ts,
      p.event || '',
      p.label || '',
      p.value || '',
      p.device || '',
      p.dev_type || '',
      p.os || '',
      p.browser || '',
      p.lang || '',
      p.tz || '',
      p.screen || '',
      p.viewport || '',
      p.dpr || '',
      p.platform || '',
      p.cores || '',
      p.mem || '',
      p.touch || '',
      p.ref || '',
      p.utm_source || '',
      p.utm_medium || '',
      p.utm_campaign || '',
      p.unlocked || '',
      p.code || '',
      p.calc_count || '',
      p.ip || '',
      p.country || '',
      p.region || '',
      p.city || '',
      p.org || ''
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
