// ============================================================
//  Учёт Поставок — Telegram Bot + WebApp Backend
//  Google Apps Script
//  Версия 2.0 — кнопки, роли, личный кабинет контрагента
// ============================================================

// ── НАСТРОЙКИ (заполнить!) ────────────────────────────────────
const SHEET_ID   = 'ВСТАВЬТЕ_ID_ТАБЛИЦЫ';       // ID Google Таблицы
const BOT_TOKEN  = 'ВСТАВЬТЕ_BOT_TOKEN';         // Токен из BotFather
const ADMIN_IDS  = [];  // Telegram user ID администраторов, напр: [123456789, 987654321]

// Ссылки на WebApp
const WEBAPP_FORM_URL = 'https://donaken12.github.io/telegram-delivery-app/';
const WEBAPP_PAY_URL  = 'https://donaken12.github.io/telegram-delivery-app/payment.html';

// Названия листов
const SH_PRODUCTS = 'Товары';
const SH_ORDERS   = 'Заказы';
const SH_PAYMENTS = 'Оплаты';
const SH_USERS    = 'Пользователи';

// ============================================================
// WEBHOOL ROUTING
// ============================================================
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    // Telegram update имеет update_id или message/callback_query
    if (body.update_id !== undefined || body.message || body.callback_query) {
      handleTelegramUpdate(body);
    } else {
      // Прямой POST из WebApp (fetch из index.html / payment.html)
      handleWebAppPost(body);
    }
  } catch (err) {
    Logger.log('doPost error: ' + err + '\n' + err.stack);
  }
  return ContentService.createTextOutput('OK');
}

// GET — список товаров для формы WebApp
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SH_PRODUCTS);
    if (!sheet) return jsonResp({ error: 'Лист "' + SH_PRODUCTS + '" не найден' });
    const rows = sheet.getDataRange().getValues();
    const products = rows.slice(1)
      .filter(r => r[0] && r[1])
      .map(r => ({ id: String(r[0]).trim(), name: String(r[1]).trim(), photo: r[2] ? String(r[2]).trim() : '' }));
    return jsonResp(products);
  } catch (err) {
    return jsonResp({ error: err.message });
  }
}

// ============================================================
// TELEGRAM HANDLERS
// ============================================================
function handleTelegramUpdate(update) {
  if (update.message)        handleMessage(update.message);
  else if (update.callback_query) handleCallbackQuery(update.callback_query);
}

function handleMessage(msg) {
  const chatId    = msg.chat.id;
  const userId    = msg.from.id;
  const text      = (msg.text || '').trim();
  const firstName = msg.from.first_name || '';
  const username  = msg.from.username  || '';

  // Данные от WebApp (tg.sendData)
  if (msg.web_app_data) {
    handleWebAppDataMsg(msg);
    return;
  }

  // Получить или создать пользователя
  const user = getOrCreateUser(userId, firstName, username);

  // Многошаговые состояния
  const state = getUserState(userId);
  if (state === 'AWAITING_DELIVERY') {
    processDeliveryTemplate(chatId, userId, text, user);
    return;
  }

  // Кнопки и команды
  switch (text) {
    case '/start':
    case '🏠 Меню':
      sendMainMenu(chatId, userId, firstName); break;

    case '📋 Отправить поставку':
    case '/template':
      sendDeliveryTemplate(chatId, userId); break;

    case '💰 Мои расчёты':
    case '/payment':
    case '/debts':
      sendContractorBalance(chatId, userId, user); break;

    case '📦 Все поставки':
    case '/status':
      sendAllDeliveries(chatId, userId); break;

    case '💳 Долги контрагентов':
      sendDebts(chatId, userId); break;

    case '📊 Статистика':
    case '/stats':
      sendStats(chatId, userId); break;

    case '👥 Пользователи':
    case '/users':
      sendUserList(chatId, userId); break;

    case '/help':
      sendHelp(chatId, userId); break;

    default:
      if (text.startsWith('/adduser')) {
        addUserCommand(chatId, userId, text);
      } else if (text.startsWith('/makeadmin')) {
        makeAdminCommand(chatId, userId, text);
      } else {
        sendMainMenu(chatId, userId, firstName);
      }
  }
}

function handleCallbackQuery(cq) {
  const chatId = cq.message.chat.id;
  const userId = cq.from.id;
  const data   = cq.data || '';
  answerCallbackQuery(cq.id);

  if (data.startsWith('pay_confirm_')) {
    const amount = parseFloat(data.replace('pay_confirm_', '')) || 0;
    recordPaymentConfirmation(chatId, userId, amount);
  } else if (data === 'cancel') {
    clearUserState(userId);
    sendMainMenu(chatId, userId, cq.from.first_name);
  } else if (data === 'my_deliveries') {
    sendMyDeliveries(chatId, userId);
  }
}

// ============================================================
// ВЛАВНОЕ МЕНЮ (зависит \ от роли)
// ============================================================
function sendMainMenu(chatId, userId, name) {
  const admin = isAdmin(userId);

  let text, keyboard;

  if (admin) {
    text = '👑 *Привет, босс!*\n\lВыберите действие:';
    keyboard = {
      keyboard: [
        [
          { text: '📓Уорма pоставки', web_app: { url: WEBAPP_FORM_URL } },
          { text: '💰 Внести оплату',  web_app: { url: WEBAPP_PAY_URL  } }
        ],
        [
          { text: '📦 Все поставки' },
          { text: '💳 Долги контрагентов' }
        ],
        [
          { text: '📊 Статистика' },
          { text: '👥 Пользователи' }
        ]
      ],
      resize_keyboard: true,
      persistent: true
    };
  } else {
    const greeting = '👋 Привет, ' + (name || 'друг') + '!\n\nЧто хотите сделать?';
    text = greeting;
    keyboard = {
      keyboard: [
        [{ text: '📋 Отправить поставку' }],
        [{ text: '💰 Мои расчёты' }]
      ],
      resize_keyboard: true,
      persistent: true
    };
  }

  sendMessage(chatId, text, keyboard, 'Markdown');
}

// ============================================================
// ШАБЛОН ПОСТАВКИ (для контрагентов)
// ============================================================
function sendDeliveryTemplate(chatId, userId) {
  const template =
    '📋 *ШАБЛОН ПОСТАВКИ*\n\n' +
    'Скопируйте, заполните и отправьте мне:\n\n' +
    '```\n' +
    'Контрагент: \n' +
    'Товар: \n' +
    'Артикул: \n' +
    'Цвет: \n' +
    'Размеры: \n' +
    'Кол-во: \n' +
    'Цена сом: \n' +
    'Полная хена: \n' +
    'Курс: \n' +
    'Адрес доставки: \n' +
    'Накладная: \n' +
    'Комментарий: \n' +
    '```\n\n' +
    '✏️ *Обязательные:* Товар, Кол-во, Цена сом\n' +
    '💡 Цена сом = по документам, Полная цена = фактическая';

  setUserState(userId, 'AWAITING_DELIVERY');

  const cancelKb = {
    keyboard: [[{ text: '❌ Отмена' }]],
    resize_keyboard: true
  };

  sendMessage(chatId, template, cancelKb, 'Markdown');
}

function processDeliveryTemplate(chatId, userId, text, user) {
  if (text === '❌ Отмена' || text === '/start') {
    clearUserState(userId);
    sendMainMenu(chatId, userId, user ? user.name : '');
    return;
  }

  // Парсим поля формата "Ключ: Значение"
  const fields = {};
  text.split('\n').forEach(function(line) {
    const colon = line.indexOf(':');
    if (colon > 0) {
      const key = line.substring(0, colon).trim().toLowerCase();
      const val = line.substring(colon + 1).trim();
      if (val) fields[key] = val;
    }
  });

  const товар  = fields['товар']  || fields['товара'] || '';
  const колво  = fields['кол-во'] || fields['количество'] || fields['кол'] || fields['кол.'] || '';
  const цена   = fields['цена сом'] || fields['цена'] || '';

  if (!товар || !колво) {
    sendMessage(chatId,
      '⚠️ Не заполнены обязательные поля: *Товар*, *Кол-во*\n\nПопробуйте ещё раз или нажмите ❌ Отмена',
      null, 'Markdown');
    return;
  }

  try {
    const ss  = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SH_ORDERS);
    if (!sheet) sheet = ss.insertSheet(SH_ORDERS);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Дата', 'Контрагент', 'Telegram ID', 'Товар', 'Артикул',
        'Цвет', 'Размеры', 'Кол-во', 'Цена сом', 'Полная цена',
        'Курс', 'Адрес', 'Накладная', 'Комментарий', 'Статус'
      ]);
      sheet.setFrozenRows(1);
    }

    const now         = new Date();
    const dateStr     = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    const contractor  = fields['контрагент'] || (user ? user.name : '') || String(userId);
    const цenaNum     = parseFloat(String(цена).replace(/[^\d.]/g, '')) || 0;
    const кол         = parseInt(String(колво).replace(/[^\d]/g, '')) || 0;

    sheet.appendRow([
      dateStr, contractor, userId,
      товар,
      fields['артикул'] || '',
      fields['цвет']    || '',
      fields['размеры'] || '',
      кол, цenaNum,
      parseFloat(String(fields['полная хена'] || '').replace(/[^\d.]/g, '')) || 0,
      parseFloat(String(fields['курс']        || '').replace(/[^\d.]/g, '')) || 0,
      fields['адрес доставки'] || fields['адрес'] || '',
      fields['накладная'] || '',
      fields['комментарий'] || '',
      'Новая'
    ]);

    clearUserState(userId);

    const confirm =
      '✅ *Поставка записана!*\n\n' +
      '👤 ' + contractor   + '\n' +
      '🛍 ' + товар        + '\n' +
      '🔢 Кол-во: ' + кол  + ' шт.\n' +
      '💵 Цена: ' + цenaNum.toLocaleString() + ' сом';

    // Уведомить всех админов
    ADMIN_IDS.forEach(function(aid) {
      try { sendMessage(aid, '📦 *Новая поставка!*\n\n' + confirm, null, 'Markdown'); } catch(e) {}
    });

    sendMessage(chatId, confirm, null, 'Markdown');

    // Вернуть меню через паузу
    Utilities.sleep(600);
    sendMainMenu(chatId, userId, user ? user.name : '');

  } catch (err) {
    Logger.log('processDeliveryTemplate: ' + err);
    sendMessage(chatId, '❌ Ошибка сохранения: ' + err.message);
  }
}

// ===================================================================================================
// РАСЧЁТЫ КОНТРАГЕНТА (личный кабинет)
// ============================================================
function sendContractorBalance(chatId, userId, user) {
  try {
    const ss          = SpreadsheetApp.openById(SHEET_ID);
    const ordersSheet = ss.getSheetByName(SH_ORDERS);
    const paySheet    = ss.getSheetByName(SH_PAYMENTS);

    // Поставки этого контрагента
    let orderRows = [];
    let totalOrders = 0;

    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const rows = ordersSheet.getDataRange().getValues().slice(1);
      orderRows  = rows.filter(r => String(r[2]) === String(userId));
      orderRows.forEach(r => { totalOrders += (parseFloat(r[8]) || 0) * (parseInt(r[7]) || 0); });
    }

    // Оплаты этого контрагента
    let totalPaid = 0;
    if (paySheet && paySheet.getLastRow() > 1) {
      paySheet.getDataRange().getValues().slice(1).forEach(function(r) {
        if (String(r[2]) === String(userId)) totalPaid += parseFloat(r[3]) || 0;
      });
    }

    const balance = totalOrders - totalPaid;
    const name    = (user && user.name) || String(userId);

    let msg =
      '💰 *Расчёты: ' + name + '*\n\n' +
      '📦 Поставок: '     + orderRows.length + ' шт.\n' +
      '💵 Сумма: '        + totalOrders.toLocaleString() + ' сом\n' +
      '✅ Оплачено: '     + totalPaid.toLocaleString()  + ' сом\n\n';

    if (balance > 0) {
      msg += '🔴 *К оплате: ' + balance.toLocaleString() + ' сом*';
    } else if (balance < 0) {
      msg += '🟢 *Переплата: ' + Math.abs(balance).toLocaleString() + ' сом*';
    } else {
      msg += '✅ *Расчёт закрыт*';
    }

    // Последние 5 поставок
    if (orderRows.length > 0) {
      msg += '\n\n📋 *Последние поставки:*';
      orderRows.slice(-5).reverse().forEach(function(r) {
        const sum = (parseFloat(r[8]) || 0) * (parseInt(r[7]) || 0);
        msg += '\n• ' + (r[3] || '?') + ' — ' + (r[7] || 0) + ' шт × ' +
               (r[8] || 0).toLocaleString() + ' = ' + sum.toLocaleString() + ' сом (' + (r[0] || '') + ')';
      });
    }

    // Кнопка подтверждения, если есть долг
    let markup = {
      inline_keyboard: [[
        { text: '🔄 Обновить', callback_data: 'my_deliveries' }
      ]]
    };
    if (balance > 0) {
      markup.inline_keyboard.unshift([{
        text: '✅ Подтвердить получение ' + balance.toLocaleString() + ' сом',
        callback_data: 'pay_confirm_' + balance
      }]);
    }

    sendMessage(chatId, msg, markup, 'Markdown');

  } catch (err) {
    Logger.log('sendContractorBalance: ' + err);
    sendMessage(chatId, '❌ Ошибка: ' + err.message);
  }
}

function sendMyDeliveries(chatId, userId) {
  const user = getUser(userId);
  sendContractorBalance(chatId, userId, user);
}

function recordPaymentConfirmation(chatId, userId, amount) {
  try {
    const ss  = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SH_PAYMENTS);
    if (!sheet) sheet = ss.insertSheet(SH_PAYMENTS);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Дата', 'Контрагент', 'Telegram ID', 'Сумма', 'За что', 'Способ', 'Комментарий']);
      sheet.setFrozenRows(1);
    }

    const user    = getUser(userId);
    const now     = new Date();
    const dateStr = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    const name    = (user && user.name) || String(userId);

    sheet.appendRow([dateStr, name, userId, amount, 'Баланс', 'Telegram подтверждение', '']);

    const adminMsg =
      '💰 *Подтверждение оплаты*\n\n' +
      '👤 ' + name + '\n' +
      '💵 ' + amount.toLocaleString() + ' сом\n' +
      '📅 ' + dateStr;

    ADMIN_IDS.forEach(function(aid) {
      try { sendMessage(aid, adminMsg, null, 'Markdown'); } catch(e) {}
    });

    sendMessage(chatId, '✅ Принято! Записано: ' + amount.toLocaleString() + ' сом', null, 'Markdown');
    Utilities.sleep(400);
    sendContractorBalance(chatId, userId, user);

  } catch (err) {
    sendMessage(chatId, '❌ ' + err.message);
  }
}

// ============================================================
// ADMIN VIEWS
// ============================================================
function sendAllDeliveries(chatId, userId) {
  if (!isAdmin(userId)) { sendMessage(chatId, '⛔ Нет доступа'); return; }
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_ORDERS);
    if (!sheet || sheet.getLastRow() < 2) { sendMessage(chatId, '📭 Поставок пока нет'); return; }

    const data = sheet.getDataRange().getValues().slice(1).slice(-15).reverse();
    let msg = '📦 *Последние поставки:*\n\n';
    data.forEach(function(r, i) {
      const sum = (parseFloat(r[8]) || 0) * (parseInt(r[7]) || 0);
      msg += (i + 1) + '. ' + (r[1] || '?') + ' — ' + (r[3] || '?') +
             ' (' + (r[7] || 0) + ' шт, ' + sum.toLocaleString() + ' сом)\n';
    });
    sendMessage(chatId, msg, null, 'Markdown');
  } catch (err) { sendMessage(chatId, '❌ ' + err.message); }
}

function sendDebts(chatId, userId) {
  if (!isAdmin(userId)) { sendMessage(chatId, '⛔ Нет доступа'); return; }
  try {
    const ss          = SpreadsheetApp.openById(SHEET_ID);
    const ordersSheet = ss.getSheetByName(SH_ORDERS);
    const paySheet    = ss.getSheetByName(SH_PAYMENTS);

    if (!ordersSheet || ordersSheet.getLastRow() < 2) { sendMessage(chatId, '📭 Поставок нет'); return; }

    const totals = {}, ids = {};
    ordersSheet.getDataRange().getValues().slice(1).forEach(function(r) {
      const name = r[1] || 'Неизвестно';
      totals[name] = (totals[name] || 0) + (parseFloat(r[8]) || 0) * (parseInt(r[7]) || 0);
      if (r[2]) ids[name] = r[2];
    });

    const paid = {};
    if (paySheet && paySheet.getLastRow() > 1) {
      paySheet.getDataRange().getValues().slice(1).forEach(function(r) {
        const name = r[1] || 'Неизвестно';
        paid[name] = (paid[name] || 0) + (parseFloat(r[3]) || 0);
      });
    }

    let msg = '💳 *Долги контрагентов:*\n\n';
    let total = 0;
    let hasAny = false;
    Object.keys(totals).sort().forEach(function(name) {
      const debt = (totals[name] || 0) - (paid[name] || 0);
      if (debt > 100) {
        msg += '🔴 ' + name + ': *' + debt.toLocaleString() + ' сом*\n';
        total += debt;
        hasAny = true;
      }
    });
    if (!hasAny) msg += '✅ Все расчёты закрыты';
    else msg += '\n💵 *Итого дебиторка: ' + total.toLocaleString() + ' сом*';

    sendMessage(chatId, msg, null, 'Markdown');
  } catch (err) { sendMessage(chatId, '❌ ' + err.message); }
}

function sendStats(chatId, userId) {
  if (!isAdmin(userId)) { sendMessage(chatId, '⛔ Нет доступа'); return; }
  try {
    const ss          = SpreadsheetApp.openById(SHEET_ID);
    const ordersSheet = ss.getSheetByName(SH_ORDERS);
    const paySheet    = ss.getSheetByName(SH_PAYMENTS);

    let ordersCount = 0, totalSum = 0;
    if (ordersSheet && ordersSheet.getLastRow() > 1) {
      const rows = ordersSheet.getDataRange().getValues().slice(1);
      ordersCount = rows.length;
      rows.forEach(r => { totalSum += (parseFloat(r[8]) || 0) * (parseInt(r[7]) || 0); });
    }

    let totalPaid = 0;
    if (paySheet && paySheet.getLastRow() > 1) {
      paySheet.getDataRange().getValues().slice(1).forEach(r => { totalPaid += parseFloat(r[3]) || 0; });
    }

    sendMessage(chatId,
      '📊 *Статистика*\n\n' +
      '📦 Поставок: '     + ordersCount + '\n' +
      '💵 Сумма: '        + totalSum.toLocaleString()   + ' сом\n' +
      '✅ Оплачено: '     + totalPaid.toLocaleString()  + ' сом\n' +
      '🔴 Дебиторка: '    + (totalSum - totalPaid).toLocaleString() + ' сом',
      null, 'Markdown');
  } catch (err) { sendMessage(chatId, '❌ ' + err.message); }
}

function sendUserList(chatId, userId) {
  if (!isAdmin(userId)) { sendMessage(chatId, '⛔ Нет доступа'); return; }
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_USERS);
    if (!sheet || sheet.getLastRow() < 2) {
      sendMessage(chatId, '👥 Пользователей нет\n\nДобавьте: /adduser [id] [имя]');
      return;
    }
    let msg = '👥 *Пользователи:*\n\n';
    sheet.getDataRange().getValues().slice(1).forEach(function(r) {
      const icon = r[2] === 'admin' ? '👑' : '👤';
      msg += icon + ' ' + (r[1] || '?') + ' (ID: ' + r[0] + ')';
      if (r[3]) msg += ' @' + r[3];
      msg += '\n';
    });
    msg += '\n➕ /adduser [telegram_id] [имя]\n👑 /makeadmin [telegram_id]';
    sendMessage(chatId, msg, null, 'Markdown');
  } catch (err) { sendMessage(chatId, '❌ ' + err.message); }
}

function addUserCommand(chatId, userId, text) {
  if (!isAdmin(userId)) { sendMessage(chatId, '⛔ Нет доступа'); return; }
  const parts = text.split(' ').filter(Boolean);
  if (parts.length < 3) {
    sendMessage(chatId, '⚠️ Формат: /adduser [telegram_id] [имя]\n\nПример: /adduser 123456789 Иван Иванов');
    return;
  }
  const newId   = parts[1];
  const newName = parts.slice(2).join(' ');
  try {
    const ss  = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SH_USERS);
    if (!sheet) sheet = ss.insertSheet(SH_USERS);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Telegram ID', 'Имя', 'Роль', 'Username', 'Дата добавления']);
      sheet.setFrozenRows(1);
    }
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    sheet.appendRow([newId, newName, 'contractor', '', now]);
    sendMessage(chatId, '✅ Добавлен: ' + newName + ' (ID: ' + newId + ')');
  } catch (err) { sendMessage(chatId, '❌ ' + err.message); }
}

function makeAdminCommand(chatId, userId, text) {
  if (!isAdmin(userId)) { sendMessage(chatId, '⛔ Нет доступа'); return; }
  const parts = text.split(' ').filter(Boolean);
  if (parts.length < 2) { sendMessage(chatId, '⚠️ Формат: /makeadmin [telegram_id]'); return; }
  const targetId = parts[1];
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SH_USERS);
    if (!sheet) { sendMessage(chatId, '❌ Лист пользователей не найден'); return; }
    const rows  = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(targetId)) {
        sheet.getRange(i + 1, 3).setValue('admin');
        sendMessage(chatId, '✅ ' + (rows[i][1] || targetId) + ' теперь администратор');
        return;
      }
    }
    sendMessage(chatId, '❌ Пользователь с ID ' + targetId + ' не найден. Сначала добавьте через /adduser');
  } catch (err) { sendMessage(chatId, '❌ ' + err.message); }
}

function sendHelp(chatId, userId) {
  const admin = isAdmin(userId);
  const msg = admin
    ? '📖 *Команды администратора:*\n\n' +
      '/adduser [id] [имя] — добавить пользователя\n' +
      '/makeadmin [id] — сделать администратор\n' +
      '/users — список пользователей\n' +
      '/stats — статистика\n' +
      '/status — все поставки\n' +
      '/debts — долги'
    : '📖 *справка�:*\n\n' +
      '📋 *Отправить поставку* — получить шаблон, заполнить и отправить\n' +
      '💰 *Мои расчёты* — ваши поставки, баланс и подтверждение оплаты';
  sendMessage(chatId, msg, null, 'Markdown');
}

// ============================================================
// WEBAPP DATA (приходит через tg.sendData из WebApp)
// ============================================================
function handleWebAppDataMsg(msg) {
  const chatId = msg.chat.id;
  const userId = msg.from.id;
  try {
    const data = JSON.parse(msg.web_app_data.data);
    if (data.type === 'webapp_delivery') {
      handleWebAppPost(data);
      const n = (data.models || []).length;
      sendMessage(chatId,
        '✅ *Заявка на поставку принята!*\n📦 Моделей: ' + n + '\n💵 Итого: ' + (data.totalSum || 0).toLocaleString() + ' руб.',
        null, 'Markdown');
    } else if (data.type === 'webapp_payment') {
      handleWebAppPost(data);
      sendMessage(chatId,
        '✅ *Оплата записана!*\n💰 ' + (data.amount || 0).toLocaleString() + ' руб. — ' + (data.purpose || ''),
        null, 'Markdown');
    }
    Utilities.sleep(400);
    sendMainMenu(chatId, userId, msg.from.first_name);
  } catch (err) {
    Logger.log('handleWebAppDataMsg: ' + err);
  }
}

// Прямой fetch-запрос из WebApp
function handleWebAppPost(data) {
  try {
    const ss  = SpreadsheetApp.openById(SHEET_ID);
    const now = new Date();
    const dt  = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');

    if (data.type === 'webapp_payment') {
      let sh = ss.getSheetByName(SH_PAYMENTS);
      if (!sh) sh = ss.insertSheet(SH_PAYMENTS);
      if (sh.getLastRow() === 0) {
        sh.appendRow(['Дата', 'Контрагент', 'Telegram ID', 'Сумма', 'За что', 'Способ', 'Комментарий']);
        sh.setFrozenRows(1);
      }
      sh.appendRow([dt, data.contractor || '', data.contractorId || '', data.amount || 0,
                    data.purpose || '', data.payMethod || '', data.comment || '']);
    } else {
      // webapp_delivery
      let sh = ss.getSheetByName(SH_ORDERS);
      if (!sh) sh = ss.insertSheet(SH_ORDERS);
      if (sh.getLastRow() === 0) {
        sh.appendRow(['Дата', 'Контрагент', 'Telegram ID', 'Дата отгрузки',
                      'Артикул', 'Товар', 'Цвет', 'Размер от', 'Размер до',
                      'Кол-во', 'Линеек', 'Цена', 'Сумма', 'Комментарий', 'Статус']);
        sh.setFrozenRows(1);
      }
      (data.models || []).forEach(function(m) {
        (m.colors || []).forEach(function(c) {
          sh.appendRow([dt, data.contractor || '', data.contractorId || '', data.shipDate || '',
                        m.productId || '', m.productName || '', c.color || '',
                        m.sizeFrom || '', m.sizeTo || '', c.total || 0,
                        c.lineupCount || '', m.price || 0,
                        (c.total || 0) * (m.price || 0), data.comment || '', 'Новая']);
        });
      });
    }
    return jsonResp({ success: true });
  } catch (err) {
    Logger.log('handleWebAppPost: ' + err);
    return jsonResp({ error: err.message });
  }
}

// ============================================================
// УПРАВЛЕНИЕ ПОЛЬЗОВАТЕЛЯМИ
// ============================================================
function getUser(telegramId) {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SH_USERS);
    if (!sheet || sheet.getLastRow() < 2) return null;
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(telegramId)) {
        return { id: rows[i][0], name: rows[i][1], role: rows[i][2] || 'contractor' };
      }
    }
    return null;
  } catch (e) { return null; }
}

function getOrCreateUser(telegramId, firstName, username) {
  let user = getUser(telegramId);
  if (user) return user;
  try {
    const ss  = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SH_USERS);
    if (!sheet) sheet = ss.insertSheet(SH_USERS);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['Telegram ID', 'Имя', 'Роль', 'Username', 'Дата добавления']);
      sheet.setFrozenRows(1);
    }
    // Первый пользователь и если в ADMIN_IDS → admin
    const inAdminList = ADMIN_IDS.indexOf(Number(telegramId)) >= 0 || ADMIN_IDS.indexOf(String(telegramId)) >= 0;
    const isFirst     = sheet.getLastRow() === 1;
    const role        = (inAdminList || isFirst) ? 'admin' : 'contractor';
    const now         = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
    const name        = firstName || username || String(telegramId);
    sheet.appendRow([telegramId, name, role, username || '', now]);
    return { id: telegramId, name: name, role: role };
  } catch (e) {
    return { id: telegramId, name: firstName || String(telegramId), role: 'contractor' };
  }
}

function isAdmin(telegramId) {
  if (ADMIN_IDS.indexOf(Number(telegramId)) >= 0 || ADMIN_IDS.indexOf(String(telegramId)) >= 0) return true;
  const u = getUser(telegramId);
  return u && u.role === 'admin';
}

// ============================================================
// СОСТОЯНИЯ (многошаговые диалоги)
// ============================================================
function getUserState(userId) {
  return PropertiesService.getUserProperties().getProperty('s_' + userId) || '';
}
function setUserState(userId, state) {
  PropertiesService.getUserProperties().setProperty('s_' + userId, state);
}
function clearUserState(userId) {
  PropertiesService.getUserProperties().deleteProperty('s_' + userId);
}

// ============================================================
// TELEGRAM API
// ============================================================
function sendMessage(chatId, text, replyMarkup, parseMode) {
  const payload = { chat_id: chatId, text: text };
  if (parseMode)   payload.parse_mode   = parseMode;
  if (replyMarkup) payload.reply_markup = JSON.stringify(replyMarkup);
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/sendMessage', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify(payload), muteHttpExceptions: true
    });
  } catch (e) { Logger.log('sendMessage err: ' + e); }
}

function answerCallbackQuery(id, text) {
  try {
    UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/answerCallbackQuery', {
      method: 'post', contentType: 'application/json',
      payload: JSON.stringify({ callback_query_id: id, text: text || '' }),
      muteHttpExceptions: true
    });
  } catch (e) {}
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// НАСТРОЙКА (запускать вручную из редактора!)
// ============================================================

/** Шаг 1. Зарегистрировать webhook */
function setWebhook() {
  const url  = ScriptApp.getService().getUrl();
  const resp = UrlFetchApp.fetch(
    'https://api.telegram.org/bot' + BOT_TOKEN + '/setWebhook?url=' + encodeURIComponent(url)
  );
  Logger.log(resp.getContentText());
}

/** Шаг 2. Установить команды в меню бота */
function setBotCommands() {
  UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/setMyCommands', {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ commands: [
      { command: 'start',    description: 'Главное меню' },
      { command: 'template', description: 'Шаблон поставки' },
      { command: 'payment',  description: 'Мои расчёты' },
      { command: 'help',     description: 'Справка' }
    ]})
  });
  Logger.log('Commands set');
}

/** Шаг 3. Сделать себя администратором (первый раз) */
function makeMyself_Admin() {
  // Вставьте ваш Telegram ID ниже:
  const MY_ID = 'ВСТАВЬТЕ_СВОЙ_TELEGRAM_ID';
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  let sheet   = ss.getSheetByName(SH_USERS);
  if (!sheet) { sheet = ss.insertSheet(SH_USERS); }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['Telegram ID', 'Имя', 'Роль', 'Username', 'Дата добавления']);
  }
  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(MY_ID)) {
      sheet.getRange(i + 1, 3).setValue('admin');
      Logger.log('Updated to admin: ' + MY_ID);
      return;
    }
  }
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy');
  sheet.appendRow([MY_ID, 'Admin', 'admin', '', now]);
  Logger.log('Added as admin: ' + MY_ID);
}

function deleteWebhook() {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/deleteWebhook');
  Logger.log(r.getContentText());
}

function getWebhookInfo() {
  const r = UrlFetchApp.fetch('https://api.telegram.org/bot' + BOT_TOKEN + '/getWebhookInfo');
  Logger.log(r.getContentText());
}
