// ============================================================
//  Google Apps Script для приложения "Поставка одежды"
//  Инструкция по установке — внизу файла
// ============================================================

// ▼ ВСТАВЬТЕ ID вашей Google Таблицы
const SHEET_ID = 'ВСТАВЬТЕ_ID_ТАБЛИЦЫ_СЮДА';

// Названия листов (можно переименовать)
const SHEET_PRODUCTS = 'Товары';   // список товаров для формы
const SHEET_ORDERS   = 'Заказы';   // куда пишутся заказы

// ============================================================
//  GET — возвращает список товаров (для дропдауна в форме)
// ============================================================
function doGet(e) {
  try {
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_PRODUCTS);

    if (!sheet) {
      return json({ error: 'Лист "' + SHEET_PRODUCTS + '" не найден' });
    }

    const rows = sheet.getDataRange().getValues();
    // Ожидаемые колонки: A=Артикул, B=Название, C=URL фото (необязательно)
    const products = rows.slice(1)          // пропускаем заголовок
      .filter(r => r[0] && r[1])            // только строки с артикулом и названием
      .map(r => ({
        id:    String(r[0]).trim(),
        name:  String(r[1]).trim(),
        photo: r[2] ? String(r[2]).trim() : ''
      }));

    return json(products);
  } catch (err) {
    return json({ error: err.message });
  }
}

// ============================================================
//  POST — записывает данные заказа в лист "Заказы"
// ============================================================
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss   = SpreadsheetApp.openById(SHEET_ID);

    let sheet = ss.getSheetByName(SHEET_ORDERS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_ORDERS);
    }

    // Создаём заголовки если лист пустой
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Дата',
        'Контрагент',
        'ID / @username',
        'Дата отгрузки',
        'Артикул',
        'Товар',
        'Цвет',
        'Размер от',
        'Размер до',
        'Кол-во',
        'Линеек',
        'Цена закупки',
        'Сумма',
        'Комментарий'
      ]);
      sheet.setFrozenRows(1);
    }

    const now        = new Date();
    const dateStr    = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    const shipDate   = data.shipDate   || '';
    const contractor = data.contractor || '';
    const contractorId = data.contractorId || '';
    const comment    = data.comment    || '';

    const rows = [];

    (data.models || []).forEach(function(model) {
      (model.colors || []).forEach(function(c) {
        rows.push([
          dateStr,
          contractor,
          contractorId,
          shipDate,
          model.productId   || '',
          model.productName || '',
          c.color           || '',
          model.sizeFrom    || '',
          model.sizeTo      || '',
          c.total           || 0,
          c.lineupCount     || '',
          model.price       || 0,
          (c.total || 0) * (model.price || 0),
          comment
        ]);
      });
    });

    if (rows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
           .setValues(rows);
    }

    return json({ success: true, rows: rows.length });

  } catch (err) {
    return json({ error: err.message });
  }
}

// ---- helpers ----
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


// ============================================================
//  ИНСТРУКЦИЯ ПО УСТАНОВКЕ
// ============================================================
//
//  1. Откройте Google Таблицу (или создайте новую).
//     Вставьте её ID из адресной строки вместо 'ВСТАВЬТЕ_ID_ТАБЛИЦЫ_СЮДА' выше.
//     ID — это длинная строка между /d/ и /edit в URL:
//     https://docs.google.com/spreadsheets/d/ВОТ_ЭТО_ID/edit
//
//  2. В таблице создайте лист с названием "Товары" и заполните его:
//     A1: Артикул   B1: Название   C1: Фото (URL)
//     A2: ART-001   B2: Куртка зимняя   C2: https://...
//     A3: ART-002   B3: Джинсы классика   C3: (пусто — фото необязательно)
//
//  3. В Таблице откройте Расширения → Apps Script.
//     Скопируйте весь этот файл в редактор Apps Script (замените содержимое).
//     Нажмите "Сохранить".
//
//  4. Нажмите "Развернуть" → "Создать новое развёртывание":
//     - Тип: Веб-приложение
//     - Выполнять как: Я (your@email.com)
//     - Кто имеет доступ: Все (Anyone)
//     Нажмите "Развернуть". Скопируйте URL веб-приложения.
//
//  5. Вставьте URL в переменную APPS_SCRIPT_URL в index.html:
//     const APPS_SCRIPT_URL = 'https://script.google.com/macros/s/ВАША_ССЫЛКА/exec';
//
//  Готово! Теперь форма будет загружать товары из листа "Товары"
//  и сохранять заказы в лист "Заказы" при каждой отправке.
//
// ============================================================
