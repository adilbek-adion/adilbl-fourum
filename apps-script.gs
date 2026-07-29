/**
 * ============================================================
 *  ПРИЁМ ЗАЯВОК С САЙТА adion form
 *  Пишет заявку в Google Таблицу и присылает письмо на Gmail.
 *
 *  Как подключить — см. файл НАСТРОЙКА-ФОРМЫ.md
 * ============================================================
 */

/* ----- НАСТРОЙКИ: поменяйте под себя ----- */
var NOTIFY_EMAIL = 'adilbek.abdrassk@gmail.com';  // куда присылать уведомления
var SHEET_NAME   = 'Заявки';                      // название листа в таблице
var SITE_NAME    = 'adion form';                  // как подписывать письма

/* Заголовки таблицы — создаются автоматически при первой заявке */
var HEADERS = ['Дата и время', 'Имя', 'Телефон', 'Почта', 'Задача', 'Страна', 'Комментарий', 'Источник', 'Страница'];


/**
 * Принимает POST-запрос от формы на сайте.
 */
function doPost(e) {
  try {
    var data = parseRequest(e);

    /* ловушка для спам-ботов: люди это поле не видят и не заполняют */
    if (data.company) {
      return json({ ok: true, skipped: 'spam' });
    }

    if (!data.name || !data.phone) {
      return json({ ok: false, error: 'Не хватает имени или телефона' });
    }

    var row = saveToSheet(data);
    sendNotification(data, row);

    return json({ ok: true, row: row });

  } catch (err) {
    /* если что-то сломалось — присылаем себе письмо с ошибкой, чтобы заявка не потерялась */
    try {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: '⚠ Ошибка приёма заявки с сайта ' + SITE_NAME,
        body: 'Текст ошибки:\n' + err + '\n\nЧто пришло:\n' + (e && e.postData ? e.postData.contents : 'нет данных')
      });
    } catch (mailErr) {}

    return json({ ok: false, error: String(err) });
  }
}


/**
 * Проверка, что скрипт опубликован: откройте адрес веб-приложения в браузере.
 */
function doGet() {
  return json({ ok: true, status: 'Приём заявок работает', time: new Date().toISOString() });
}


/* ============================================================
   ВСПОМОГАТЕЛЬНОЕ
   ============================================================ */

/** Разбирает тело запроса: JSON или обычная форма. */
function parseRequest(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      return JSON.parse(e.postData.contents);
    } catch (err) {
      /* не JSON — пробуем как форму */
    }
  }
  return (e && e.parameter) ? e.parameter : {};
}


/** Записывает заявку в таблицу и возвращает номер строки. */
function saveToSheet(data) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  /* первая заявка — рисуем шапку */
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    var head = sheet.getRange(1, 1, 1, HEADERS.length);
    head.setFontWeight('bold').setBackground('#12294f').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);
    sheet.setColumnWidth(2, 140);
    sheet.setColumnWidth(3, 150);
    sheet.setColumnWidth(4, 220);
    sheet.setColumnWidth(7, 300);
  }

  sheet.appendRow([
    new Date(),
    data.name    || '',
    data.phone   || '',
    data.email   || '',
    data.task    || '',
    data.country || '',
    data.comment || '',
    data.source  || 'сайт',
    data.page    || ''
  ]);

  return sheet.getLastRow();
}


/** Присылает письмо на Gmail. Ответить клиенту можно прямо из письма. */
function sendNotification(data, row) {
  var subject = 'Заявка с сайта: ' + (data.task || 'проект') + ' — ' + data.name;

  var html =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:560px">' +
      '<div style="background:#12294f;color:#fff;padding:20px 24px;border-radius:12px 12px 0 0">' +
        '<div style="font-size:12px;letter-spacing:.14em;text-transform:uppercase;opacity:.7">' + SITE_NAME + '</div>' +
        '<div style="font-size:22px;font-weight:bold;margin-top:4px">Новая заявка с сайта</div>' +
      '</div>' +
      '<table style="width:100%;border-collapse:collapse;border:1px solid #e3e8ef;border-top:none">' +
        rowHtml('Имя', data.name) +
        rowHtml('Телефон', linkTel(data.phone)) +
        rowHtml('Почта', linkMail(data.email)) +
        rowHtml('Задача', data.task) +
        rowHtml('Страна', data.country) +
        rowHtml('Комментарий', data.comment) +
        rowHtml('Страница', data.page) +
        rowHtml('Время', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm')) +
      '</table>' +
      '<p style="color:#6b7690;font-size:13px;margin-top:14px">' +
        'Заявка записана в таблицу, строка ' + row + '. Чтобы ответить клиенту — просто нажмите «Ответить» в этом письме.' +
      '</p>' +
    '</div>';

  var options = {
    to: NOTIFY_EMAIL,
    subject: subject,
    htmlBody: html,
    name: SITE_NAME
  };

  /* ответ на письмо уйдёт клиенту, а не самому себе */
  if (data.email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(data.email)) {
    options.replyTo = data.email;
  }

  MailApp.sendEmail(options);
}


function rowHtml(label, value) {
  if (!value) return '';
  return '<tr>' +
    '<td style="padding:11px 24px;border-bottom:1px solid #eef1f6;color:#6b7690;width:130px;font-size:14px">' + label + '</td>' +
    '<td style="padding:11px 24px;border-bottom:1px solid #eef1f6;color:#141922;font-size:15px;font-weight:bold">' + value + '</td>' +
  '</tr>';
}

function linkTel(v) {
  return v ? '<a href="tel:' + String(v).replace(/[^\d+]/g, '') + '" style="color:#1668c9;text-decoration:none">' + v + '</a>' : '';
}

function linkMail(v) {
  return v ? '<a href="mailto:' + v + '" style="color:#1668c9;text-decoration:none">' + v + '</a>' : '';
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/**
 * Тестовая заявка — запустите эту функцию в редакторе Apps Script,
 * чтобы проверить и таблицу, и письмо, не трогая сайт.
 */
function testLead() {
  var demo = {
    name: 'Тестовый клиент',
    phone: '+7 700 123 45 67',
    email: 'test@example.com',
    task: 'лендинг',
    country: 'Казахстан',
    comment: 'Это тестовая заявка, отправлена из редактора скриптов.',
    source: 'тест',
    page: 'index.html'
  };
  var row = saveToSheet(demo);
  sendNotification(demo, row);
  Logger.log('Готово: записана строка ' + row + ', письмо отправлено на ' + NOTIFY_EMAIL);
}
