// ============================================================
// УЧЁТНАЯ СИСТЕМА ОПЛАТ ТЕЛЕМАТИКИ IGETIS
// Google Apps Script
// ============================================================

// ---- НАСТРОЙКИ (заполни перед запуском) ----
var IGETIS_TOKEN = PropertiesService.getScriptProperties().getProperty('IGETIS_TOKEN') || '';
var NOTIFY_EMAIL = 'we@igetis.pro';
var NOTIFY_DAYS  = 30;

// ---- НАЗВАНИЯ ЛИСТОВ ----
var SHEET_VEHICLES  = '🚛 Автомобили';
var SHEET_PAYERS    = '🏢 Плательщики';
var SHEET_INVOICES  = '🧾 Счета';
var SHEET_DASHBOARD = '📊 Дашборд';

// ============================================================
// ВСПОМОГАТЕЛЬНАЯ: очищает битые значения из массива строки
// getValues() возвращает объект Error для ячеек с #ERROR! формулами
// ============================================================
function safeVal(v) {
  if (v instanceof Error) return '';
  if (typeof v === 'string' && v.charAt(0) === '#') return '';
  return v;
}

// ============================================================
// ИНИЦИАЛИЗАЦИЯ: создаёт меню при открытии таблицы
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🚛 Телематика')
    .addItem('🔍 Поиск по VIN / трекеру', 'openSearchSidebar')
    .addItem('🧾 Создать счёт...', 'openInvoiceSidebar')
    .addSeparator()
    .addItem('Синхронизировать авто из API', 'syncVehiclesFromAPI')
    .addItem('🏢 Синхронизировать плательщиков из API', 'syncPayersFromAPI')
    .addItem('🔍 Диагностика API', 'debugAPI')
    .addItem('🔎 Найти устройство по VIN в API', 'debugVIN')
    .addItem('🔬 Диагностика последнего импорта', 'debugImportFile')
    .addSeparator()
    .addItem('Обновить дашборд', 'updateDashboard')
    .addItem('Проверить и отправить уведомления', 'checkAndNotify')
    .addSeparator()
    .addItem('🔄 Обновить список плательщиков в dropdown', 'refreshPayerDropdowns')
    .addItem('📅 Обновить «Подключен до» из счетов', 'refreshConnectedDates')
    .addItem('📥 Импорт из файла / добавить вручную', 'openImportSidebar')
    .addItem('✅ Аудит базы данных', 'runAudit')
    .addSeparator()
    .addItem('⚙️ Первичная настройка листов', 'setupSheets')
    .addItem('🔧 Обновить схему таблицы (добавить новые столбцы)', 'migrateSchema')
    .addItem('🔄 Восстановить пропавшие авто', 'recoverLostVehicles')
    .addToUi();
}

// ============================================================
// ПЕРВИЧНАЯ НАСТРОЙКА: создаёт заголовки на всех листах
// ============================================================
function setupSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- Плательщики ---
  var shPayers = getOrCreateSheet(ss, SHEET_PAYERS);
  if (shPayers.getLastRow() === 0) {
    shPayers.getRange(1, 1, 1, 6).setValues([['ID', 'Название юрлица', 'Контактное лицо', 'Телефон', 'Email', 'Номер договора']]);
    formatHeader(shPayers, 6);
    shPayers.setColumnWidth(1, 280);
    shPayers.setColumnWidth(2, 220);
    shPayers.setColumnWidth(6, 140);
  }

  // --- Автомобили (11 столбцов) ---
  var shVeh = getOrCreateSheet(ss, SHEET_VEHICLES);
  if (shVeh.getLastRow() === 0) {
    shVeh.getRange(1, 1, 1, 11).setValues([[
      'ID Igetis', 'VIN', 'Марка / Модель', 'Гос. номер', 'Номер трекера',
      'Дата монтажа', 'Дата подключения', 'Плательщик',
      'Эксплуатант', 'Комментарий', 'ID Плательщика (contragentId)', 'Подключен до'
    ]]);
    formatHeader(shVeh, 11);
    shVeh.setColumnWidth(2, 160);
    shVeh.setColumnWidth(3, 180);
    shVeh.setColumnWidth(5, 130);
    shVeh.setColumnWidth(8, 220);
    shVeh.setColumnWidth(11, 280);
    // Столбец K — серый (справочный, не для редактирования)
    shVeh.getRange(1, 11).setFontColor('#888888');
    shVeh.getRange(1, 11).setNote('Заполняется автоматически из API. Используй для определения плательщика.');
  }

  // --- Счета ---
  var shInv = getOrCreateSheet(ss, SHEET_INVOICES);
  if (shInv.getLastRow() === 0) {
    shInv.getRange(1, 1, 1, 10).setValues([[
      '№ счета', 'ID Плательщика', 'Название плательщика',
      'Список VIN', 'Период с', 'Период по',
      'Сумма (руб.)', 'Дата выставления', 'Дата оплаты', 'Статус'
    ]]);
    formatHeader(shInv, 10);
    shInv.setColumnWidth(4, 250);
    shInv.setColumnWidth(3, 200);
  }

  // --- Дашборд ---
  var shDash = getOrCreateSheet(ss, SHEET_DASHBOARD);
  shDash.clearContents();
  buildDashboard(shDash);

  SpreadsheetApp.getUi().alert('✅ Листы настроены!\n\nДальше:\n1. Синхронизировать плательщиков из API\n2. Вписать названия организаций в жёлтые строки\n3. Синхронизировать авто из API');
}

// ============================================================
// ДИАГНОСТИКА API
// ============================================================
function debugAPI() {
  var url = 'https://api.lk.igetis.pro/v1/devices/my';
  var response = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + IGETIS_TOKEN },
    muteHttpExceptions: true
  });
  var code = response.getResponseCode();
  var text = response.getContentText();
  var preview = text.length > 3000 ? text.substring(0, 3000) + '\n...(обрезано)' : text;
  SpreadsheetApp.getUi().alert('HTTP ' + code + '\n\nОТВЕТ API:\n' + preview);
  Logger.log('HTTP ' + code + '\n' + text);
}

// ============================================================
// СИНХРОНИЗАЦИЯ АВТО ИЗ API
// Столбцы A-E, K — из API; F,G,H,I,J — ручные (не перезаписываются)
// ============================================================
function syncVehiclesFromAPI() {
  if (IGETIS_TOKEN === 'ВСТАВЬ_ТОКЕН_СЮДА') {
    SpreadsheetApp.getUi().alert('❌ Сначала вставь токен Igetis в переменную IGETIS_TOKEN!');
    return;
  }

  var options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + IGETIS_TOKEN },
    muteHttpExceptions: true
  };

  try {
    var response = UrlFetchApp.fetch('https://api.lk.igetis.pro/v1/devices/my?IsShowDisabled=true', options);
    if (response.getResponseCode() !== 200) {
      SpreadsheetApp.getUi().alert('❌ Ошибка API: код ' + response.getResponseCode() + '\n' + response.getContentText());
      return;
    }

    var data = JSON.parse(response.getContentText());
    var devices = data.results || data.devices || data.result || (Array.isArray(data) ? data : []);

    if (!Array.isArray(devices) || devices.length === 0) {
      SpreadsheetApp.getUi().alert('⚠️ API вернул пустой список. Проверь токен и права доступа.');
      return;
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_VEHICLES);
    if (!sh) {
      SpreadsheetApp.getUi().alert('❌ Лист "' + SHEET_VEHICLES + '" не найден. Запусти "Первичную настройку".');
      return;
    }

    // Читаем существующие данные — сохраняем ручные поля, очищаем ошибки
    var existingData = {};
    var lastRow = sh.getLastRow();
    if (lastRow > 1) {
      sh.getRange(2, 1, lastRow - 1, 12).getValues().forEach(function(row) {
        var id = String(safeVal(row[0]));
        if (id) existingData[id] = row.map(safeVal);
      });
    }

    // Удаляем старые столбцы с формулами (точные совпадения названий старых версий)
    var staleHeaders = ['Название плательщика', 'ID Плательщика'];
    var headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    for (var col = headerRow.length; col >= 1; col--) {
      // Удаляем только точное совпадение (не наш новый столбец)
      if (staleHeaders.indexOf(String(headerRow[col-1])) !== -1) {
        sh.deleteColumn(col);
        Logger.log('Удалён устаревший столбец: ' + headerRow[col-1]);
      }
    }

    // Формируем строки (12 столбцов), пропускаем устройства без VIN
    var skippedNoVin = 0;
    var newRows = [];
    devices.forEach(function(d) {
      var vin = String(d.vinNumber || d.vin || d.VIN || d.vinCode || '').trim();
      if (!vin) { skippedNoVin++; return; } // нет VIN — пропускаем
      var id = String(d.deviceId || d.id || '');
      var old = existingData[id] || [];
      newRows.push([
        id,                               // A: ID Igetis
        vin,                              // B: VIN
        d.carModel   || old[2] || '',     // C: Марка/Модель
        d.carNumber  || old[3] || '',     // D: Гос.номер
        d.deviceNumber || old[4] || '',   // E: Номер трекера
        old[5] || '',                     // F: Дата монтажа (ручной)
        old[6] || '',                     // G: Дата подключения (ручной)
        old[7] || '',                     // H: Плательщик (ручной/dropdown)
        old[8] || '',                     // I: Эксплуатант (ручной)
        old[9] || '',                     // J: Комментарий (ручной)
        d.contragentId || old[10] || '',  // K: ID Плательщика (из API)
        old[11] || ''                     // L: Подключен до (сохраняем)
      ]);
    });

    // Сохраняем ВСЕ строки которых нет в API — ручные, импортированные, любые
    // Критерий: ID не совпадает ни с одним deviceId из API-ответа
    var apiIds = {};
    devices.forEach(function(d) { 
      var id = String(d.deviceId || d.id || '');
      if (id) apiIds[id] = true;
    });

    Object.keys(existingData).forEach(function(id) {
      if (apiIds[id]) return;           // это API-устройство — уже в newRows
      var old = existingData[id];
      var vin = String(old[1] || '').trim();
      if (!old[0] && !vin) return;      // совсем пустая строка — пропускаем
      // Не дублируем если VIN уже есть среди API-строк
      var vinAlreadyExists = vin && newRows.some(function(r){ return String(r[1]).trim() === vin; });
      if (!vinAlreadyExists) newRows.push(old);
    });

    // Записываем
    if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, 12).clearContent();
    if (newRows.length > 0) sh.getRange(2, 1, newRows.length, 12).setValues(newRows);

    // Зелёный фон для ручных строк (не-API)
    newRows.forEach(function(row, idx) {
      if (!apiIds[String(row[0])]) {
        sh.getRange(idx + 2, 1, 1, 12).setBackground('#e6f4ea');
      }
    });

    // Dropdown для H
    applyPayerDropdown(sh, newRows.length);

    // Авто-распределение по плательщикам (заполняет H по contragentId)
    autoAssignPayers_(sh, newRows.length);

    updateDashboard();
    SpreadsheetApp.getUi().alert(
      '✅ Синхронизировано: ' + newRows.length + ' ТС.' + (skippedNoVin ? '\n⚠️ Пропущено без VIN: ' + skippedNoVin : '') + '\n\n' +
      'Ручные поля сохранены: Дата монтажа (F), Дата подключения (G),\nЭксплуатант (I), Комментарий (J).\n\n' +
      'Если плательщик не назначен — заполни столбец H вручную через dropdown.'
    );

  } catch(e) {
    SpreadsheetApp.getUi().alert('❌ Ошибка: ' + e.message);
  }
}

// ============================================================
// СИНХРОНИЗАЦИЯ ПЛАТЕЛЬЩИКОВ ИЗ API
// /v1/profile → текущий пользователь (contragent известен)
// /v1/devices/my → группировка по contragentId → остальные плательщики
// ============================================================
function syncPayersFromAPI() {
  var ui = SpreadsheetApp.getUi();
  var options = {
    method: 'get',
    headers: { 'Authorization': 'Bearer ' + IGETIS_TOKEN },
    muteHttpExceptions: true
  };

  // Профиль текущего пользователя
  var profileResp = UrlFetchApp.fetch('https://api.lk.igetis.pro/v1/profile', options);
  if (profileResp.getResponseCode() !== 200) {
    ui.alert('❌ Ошибка /v1/profile: ' + profileResp.getResponseCode() + '\n' + profileResp.getContentText());
    return;
  }
  var profile = JSON.parse(profileResp.getContentText()).result;
  Logger.log('Profile contragentId: ' + profile.contragentId);

  // Все устройства
  var devResp = UrlFetchApp.fetch('https://api.lk.igetis.pro/v1/devices/my', options);
  if (devResp.getResponseCode() !== 200) {
    ui.alert('❌ Ошибка /v1/devices/my: ' + devResp.getResponseCode());
    return;
  }
  var devices = JSON.parse(devResp.getContentText()).results || [];

  // Группируем устройства по contragentId
  var contragentMap = {};
  devices.forEach(function(d) {
    var cid = d.contragentId || '';
    if (!contragentMap[cid]) contragentMap[cid] = [];
    contragentMap[cid].push(d.deviceId);
  });

  // Строим список плательщиков
  var payers = [];
  Object.keys(contragentMap).forEach(function(cid) {
    var isCurrent = (cid === profile.contragentId);
    payers.push({
      contragentId:        cid,
      contragentCode:      isCurrent ? (profile.contragentCode || cid) : cid,
      contragentSmallName: isCurrent ? (profile.contragentSmallName || '') : '',
      lastName:            isCurrent ? (profile.lastName || '') : '',
      name:                isCurrent ? (profile.name || '') : '',
      middleName:          isCurrent ? (profile.middleName || '') : '',
      phoneNumber:         isCurrent ? (profile.phoneNumber || '') : '',
      email:               isCurrent ? (profile.email || '') : '',
      deviceIds:           contragentMap[cid],
      isCurrent:           isCurrent
    });
  });

  // Текущий пользователь — первым
  payers.sort(function(a, b) { return b.isCurrent - a.isCurrent; });

  writePayersToSheet_(payers);

  // После записи плательщиков — распределяем авто
  var shVeh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_VEHICLES);
  if (shVeh && shVeh.getLastRow() > 1) {
    autoAssignPayers_(shVeh, shVeh.getLastRow() - 1);
  }

  var unknownCount = payers.filter(function(p) { return !p.isCurrent; }).length;
  var msg = '✅ Найдено плательщиков: ' + payers.length + '\n\n';
  if (unknownCount > 0) {
    msg += '⚠️ ' + unknownCount + ' организаций без названия — заполни столбец B.\n';
    msg += 'Подсказка: наведи на жёлтую ячейку — увидишь список авто этого плательщика.\n\n';
  }
  msg += 'Авто распределены автоматически.';
  ui.alert(msg);
}

// Записываем плательщиков на лист, сохраняя ручные названия
function writePayersToSheet_(payers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PAYERS);
  if (!sh) return;

  // Читаем существующие — сохраняем вручную введённые названия
  var existingMap = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues().forEach(function(r) {
      var id = String(safeVal(r[0]));
      if (id) existingMap[id] = r.map(safeVal);
    });
  }

  // Определяем: введено ли реальное название вручную
  function hasManualName(cid, p) {
    var ex = existingMap[cid];
    if (!ex || !ex[1]) return false;
    var v = String(ex[1]);
    return v !== '' &&
           v.indexOf('введи название') === -1 &&
           v !== p.contragentCode &&
           v !== p.contragentId;
  }

  // Строим строки
  var apiRows = payers.map(function(p) {
    var ex = existingMap[p.contragentId] || [];
    var fio = [p.lastName, p.name, p.middleName].filter(Boolean).join(' ');
    // Если есть вручную введённое имя — сохраняем его; иначе берём из API или код
    var displayName = hasManualName(p.contragentId, p)
      ? ex[1]
      : (p.contragentSmallName || p.contragentCode || p.contragentId);
    return [
      p.contragentId,               // A: ID
      displayName,                  // B: Название
      fio || ex[2] || '',           // C: Контактное лицо
      p.phoneNumber || ex[3] || '', // D: Телефон
      p.email || ex[4] || '',       // E: Email
      ex[5] || ''                   // F: Номер договора (ручной)
    ];
  });

  var apiIds = payers.map(function(p) { return p.contragentId; });
  var manualRows = [];
  Object.keys(existingMap).forEach(function(id) {
    if (apiIds.indexOf(id) === -1) manualRows.push(existingMap[id]);
  });

  var allRows = apiRows.concat(manualRows);

  sh.getRange(2, 1, Math.max(sh.getLastRow(), 2), 6).clearContent().clearFormat();
  if (allRows.length === 0) return;

  sh.getRange(2, 1, allRows.length, 6).setValues(allRows);

  // Цветовая кодировка
  for (var i = 0; i < apiRows.length; i++) {
    var realName = hasManualName(payers[i].contragentId, payers[i]) || payers[i].contragentSmallName;
    sh.getRange(i + 2, 1, 1, 6).setBackground(realName ? '#e8f0fe' : '#fff3cd');
  }
  for (var j = 0; j < manualRows.length; j++) {
    sh.getRange(apiRows.length + j + 2, 1, 1, 5).setBackground('#f9f9f9');
  }

  // Легенда
  sh.getRange(1, 7)
    .setValue('🔵 из API (имя известно)   🟡 из API (заполни название)   ⚪ ручной ввод')
    .setFontColor('#888').setFontSize(9);

  // Подсказка на жёлтых ячейках — список авто этого contragentId
  var shVeh = ss.getSheetByName(SHEET_VEHICLES);
  for (var r = 0; r < apiRows.length; r++) {
    var cid = apiRows[r][0];
    var nameCell = sh.getRange(r + 2, 2);
    // Если название — это код (не реальное имя), добавляем подсказку
    var isPlaceholderName = !hasManualName(cid, payers[r]) && !payers[r].contragentSmallName;
    if (isPlaceholderName) {
      var carList = [];
      if (shVeh && shVeh.getLastRow() > 1) {
        shVeh.getRange(2, 1, shVeh.getLastRow() - 1, 11).getValues().forEach(function(vRow) {
          if (String(safeVal(vRow[10])) === String(cid) && carList.length < 7) {
            carList.push(safeVal(vRow[1]) || safeVal(vRow[0]));
          }
        });
      }
      nameCell.setNote(
        'contragentId: ' + cid + '\n' +
        'Автомобилей: ' + carList.length + (carList.length ? '\n' + carList.join('\n') : '') + '\n\n' +
        'Введи официальное название организации.'
      );
      nameCell.setFontColor('#888888').setFontStyle('italic');
    } else {
      nameCell.clearNote().setFontColor('#000000').setFontStyle('normal');
    }
  }

  // Обновляем dropdown в Автомобилях
  if (shVeh && shVeh.getLastRow() > 1) {
    applyPayerDropdown(shVeh, shVeh.getLastRow() - 1);
  }
}

// ============================================================
// АВТО-РАСПРЕДЕЛЕНИЕ ПЛАТЕЛЬЩИКОВ
// Заполняет столбец H по совпадению contragentId (столбец K)
// с листом Плательщики. Не перезаписывает корректно заполненные.
// ============================================================
function autoAssignPayers_(shVeh, count) {
  if (!shVeh || count < 1) return;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shPay = ss.getSheetByName(SHEET_PAYERS);
  if (!shPay || shPay.getLastRow() < 2) return;

  // Карта: contragentId -> displayName (из листа Плательщики)
  var cidToName = {};
  var validNames = [];
  shPay.getRange(2, 1, shPay.getLastRow() - 1, 6).getValues().forEach(function(r) {
    var id   = String(safeVal(r[0]));
    var name = String(safeVal(r[1]));
    if (id && name && name.indexOf('введи название') === -1) {
      cidToName[id] = name;
      validNames.push(name);
    }
  });

  var vehData = shVeh.getRange(2, 1, count, 12).getValues();
  var updated = 0;

  vehData.forEach(function(row, idx) {
    var cid         = String(safeVal(row[10])); // K: contragentId
    var currentH    = String(safeVal(row[7]));   // H: текущий плательщик
    var newName     = cid ? cidToName[cid] : null;
    if (!newName) return; // для этого contragentId нет плательщика с именем

    // Назначаем если H пусто, или содержит невалидное значение (ошибку / старый UUID)
    var isValid = currentH && validNames.indexOf(currentH) !== -1;
    if (!isValid) {
      shVeh.getRange(idx + 2, 8).setValue(newName);
      updated++;
    }
  });

  Logger.log('autoAssignPayers_: обновлено ' + updated + ' строк');
}

// ============================================================
// ПРОВЕРКА УВЕДОМЛЕНИЙ
// ============================================================
function checkAndNotify() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shInv = ss.getSheetByName(SHEET_INVOICES);
  var shVeh = ss.getSheetByName(SHEET_VEHICLES);
  if (!shInv || !shVeh) return;

  var today = new Date(); today.setHours(0,0,0,0);
  var notifyDate = new Date(today);
  notifyDate.setDate(notifyDate.getDate() + NOTIFY_DAYS);

  var invRows = shInv.getLastRow() > 1
    ? shInv.getRange(2, 1, shInv.getLastRow()-1, 10).getValues().map(function(r){return r.map(safeVal);})
    : [];

  var vehRows = shVeh.getLastRow() > 1
    ? shVeh.getRange(2, 1, shVeh.getLastRow()-1, 11).getValues().map(function(r){return r.map(safeVal);})
    : [];

  var needInvoice = [], overdueList = [];

  vehRows.forEach(function(veh) {
    var vin = veh[1];
    if (!vin) return;
    var payerName = String(veh[7] || '—');

    var vehicleInvoices = invRows.filter(function(inv) {
      return String(inv[3]).indexOf(vin) !== -1;
    });

    if (vehicleInvoices.length === 0) {
      needInvoice.push({ vin: vin, payer: payerName, reason: 'нет ни одного счёта' });
      return;
    }

    vehicleInvoices.sort(function(a,b){ return new Date(b[5]) - new Date(a[5]); });
    var lastInv  = vehicleInvoices[0];
    var periodTo = new Date(lastInv[5]);
    var paidDate = lastInv[8];

    if (periodTo < today && !paidDate) {
      overdueList.push({ vin: vin, payer: payerName, invoice: lastInv[0], periodTo: formatDate(periodTo), amount: lastInv[6] });
    }
    if (periodTo >= today && periodTo <= notifyDate) {
      needInvoice.push({ vin: vin, payer: payerName, invoice: lastInv[0], periodTo: formatDate(periodTo),
        reason: 'период заканчивается через ' + Math.round((periodTo-today)/86400000) + ' дн.' });
    }
  });

  if (needInvoice.length === 0 && overdueList.length === 0) {
    Logger.log('Уведомлений нет');
    return;
  }

  GmailApp.sendEmail(NOTIFY_EMAIL,
    '🔔 Телематика: требуют внимания ' + (needInvoice.length + overdueList.length) + ' ТС',
    '', { htmlBody: buildEmailBody(needInvoice, overdueList, today) });
  Logger.log('Уведомление отправлено на ' + NOTIFY_EMAIL);

  updateDashboard();
}

// ============================================================
// ДАШБОРД
// ============================================================
function updateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shDash = ss.getSheetByName(SHEET_DASHBOARD);
  if (!shDash) return;
  shDash.clearContents();
  shDash.clearFormats();
  shDash.clearNotes();
  buildDashboard(shDash);
}

function buildDashboard(sh) {
  var today = new Date(); today.setHours(0,0,0,0);
  var notifyDate = new Date(today); notifyDate.setDate(notifyDate.getDate() + NOTIFY_DAYS);

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var shInv = ss.getSheetByName(SHEET_INVOICES);
  var shVeh = ss.getSheetByName(SHEET_VEHICLES);

  var invRows = (shInv && shInv.getLastRow() > 1)
    ? shInv.getRange(2,1,shInv.getLastRow()-1,10).getValues().map(function(r){return r.map(safeVal);}) : [];
  var vehRows = (shVeh && shVeh.getLastRow() > 1)
    ? shVeh.getRange(2,1,shVeh.getLastRow()-1,12).getValues().map(function(r){return r.map(safeVal);}) : [];

  // Строим карту: VIN -> последний счёт
  var vinToLastInv = {};
  invRows.forEach(function(inv) {
    if (!inv[0]) return;
    var vins = String(inv[3]).split(/[,;]/).map(function(v){ return v.trim(); }).filter(Boolean);
    var vinCnt = vins.length || 1;
    var amountPerVin = Math.round((Number(inv[6]) || 0) / vinCnt);
    vins.forEach(function(vin) {
      var existing = vinToLastInv[vin];
      if (!existing || new Date(inv[5]) > new Date(existing.periodTo)) {
        vinToLastInv[vin] = {
          num:         inv[0],
          periodTo:    inv[5],
          paidDate:    inv[8],
          amountTotal: Number(inv[6]) || 0,
          amountPerVin: amountPerVin,
          totalVins:   vinCnt
        };
      }
    });
  });

  // Считаем метрики
  var totalVeh = 0, noPayerVeh = 0;
  var sections = {
    noInvoice:  [],   // никогда не выставляли счёт
    overdue:    [],   // счёт просрочен и не оплачен
    expiringSoon: [], // оплачен, но подключение заканчивается
    awaitPay:   [],   // счёт выставлен, ждём оплаты
    ok:         []    // всё хорошо
  };

  vehRows.forEach(function(veh) {
    if (!veh[0] && !veh[1]) return;
    totalVeh++;
    var vin       = String(veh[1]).trim();
    var model     = String(veh[2]).trim() || '—';
    var reg       = String(veh[3]).trim();
    var payer     = String(veh[7]).trim() || '—';
    var connTo    = veh[11];
    if (!veh[7]) noPayerVeh++;

    var invRec = vin ? vinToLastInv[vin] : null;
    var item = { vin:vin, model:model, reg:reg, payer:payer, connTo:connTo };

    if (!invRec) {
      sections.noInvoice.push(item);
      return;
    }

    var periodTo = new Date(invRec.periodTo);
    var paidDate = invRec.paidDate;
    var daysLeft = Math.round((periodTo - today) / 86400000);
    item.invNum      = invRec.num;
    item.amount      = invRec.amountPerVin;   // сумма за ОДНО авто
    item.totalVins   = invRec.totalVins;       // сколько авто в счёте
    item.amountTotal = invRec.amountTotal;     // полная сумма счёта

    if (!paidDate && periodTo < today) {
      item.daysOverdue = -daysLeft;
      sections.overdue.push(item);
    } else if (!paidDate) {
      item.daysLeft = daysLeft;
      sections.awaitPay.push(item);
    } else {
      var connDt = connTo ? new Date(connTo) : periodTo;
      var daysToExpiry = Math.round((connDt - today) / 86400000);
      if (daysToExpiry < 0) {
        item.daysExpired = -daysToExpiry;
        sections.expiringSoon.push(item);
      } else if (connDt <= notifyDate) {
        item.daysLeft = daysToExpiry;
        sections.expiringSoon.push(item);
      } else {
        sections.ok.push(item);
      }
    }
  });

  var totalReceivable = 0;
  sections.overdue.forEach(function(i){ totalReceivable += Number(i.amount)||0; });
  sections.awaitPay.forEach(function(i){ totalReceivable += Number(i.amount)||0; });

  // Очищаем лист
  sh.clearContents(); sh.clearFormats();
  var r = 1;

  function writeTitle(text, bg, color) {
    var cell = sh.getRange(r, 1);
    sh.getRange(r, 1, 1, 5).merge().setValue(text)
      .setBackground(bg || '#1a73e8').setFontColor(color || '#ffffff')
      .setFontWeight('bold').setFontSize(12)
      .setVerticalAlignment('middle');
    sh.setRowHeight(r, 32);
    r++;
  }

  function writeSectionHeader(icon, title, count, bg, explanation) {
    sh.getRange(r, 1, 1, 5).setBackground(bg).setBorder(false,false,false,false,false,false);
    sh.getRange(r, 1).setValue(icon + ' ' + title).setFontWeight('bold').setFontSize(11).setBackground(bg);
    sh.getRange(r, 4).setValue('Авто: ' + count).setFontWeight('bold').setBackground(bg).setHorizontalAlignment('right');
    sh.setRowHeight(r, 26);
    r++;
    if (explanation) {
      sh.getRange(r, 1, 1, 5).merge().setValue(explanation)
        .setFontColor('#5f6368').setFontStyle('italic').setFontSize(9)
        .setBackground(bg).setWrap(true);
      sh.setRowHeight(r, 28);
      r++;
    }
  }

  function writeTableHeader(cols) {
    var row = sh.getRange(r, 1, 1, cols.length);
    row.setValues([cols]).setFontWeight('bold').setBackground('#f1f3f4').setFontSize(9);
    sh.setRowHeight(r, 18);
    r++;
  }

  function writeRow(values, bg) {
    sh.getRange(r, 1, 1, values.length).setValues([values]).setBackground(bg || '#ffffff').setFontSize(9);
    sh.setRowHeight(r, 16);
    r++;
  }

  function writeEmpty() { sh.setRowHeight(r, 8); r++; }

  function fmtDate(d) {
    if (!d) return '—';
    var dt = new Date(d);
    if (isNaN(dt)) return '—';
    return ('0'+dt.getDate()).slice(-2)+'.'+('0'+(dt.getMonth()+1)).slice(-2)+'.'+dt.getFullYear();
  }

  function fmtMoney(n) {
    if (!n) return '—';
    return Number(n).toLocaleString('ru-RU') + ' ₽';
  }

  // ── Заголовок ──
  var dateStr = Utilities.formatDate(today, 'Europe/Moscow', 'dd.MM.yyyy HH:mm');
  writeTitle('📊 ДАШБОРД ТЕЛЕМАТИКИ — Igetis     Обновлено: ' + dateStr, '#1a73e8', '#ffffff');

  // ── Сводка ──
  sh.getRange(r, 1, 1, 5).setBackground('#e8f0fe');
  sh.getRange(r, 1).setValue('Всего авто в базе: ' + totalVeh).setFontWeight('bold').setBackground('#e8f0fe');
  sh.getRange(r, 2).setValue('К получению: ' + fmtMoney(totalReceivable)).setFontWeight('bold').setBackground('#e8f0fe').setFontColor(totalReceivable > 0 ? '#d93025' : '#137333');
  sh.getRange(r, 3).setValue('Оплачено счетов: ' + sections.ok.length).setBackground('#e8f0fe').setFontColor('#137333');
  sh.getRange(r, 4).setValue('Без плательщика: ' + noPayerVeh).setBackground('#e8f0fe').setFontColor(noPayerVeh > 0 ? '#d93025' : '#5f6368');
  sh.setRowHeight(r, 24); r++;
  writeEmpty();

  // ────────────────────────────────────
  // СЕКЦИЯ 1: Просроченные
  // ────────────────────────────────────
  if (sections.overdue.length > 0) {
    writeSectionHeader('🔴', 'ТРЕБУЮТ НЕМЕДЛЕННЫХ ДЕЙСТВИЙ — счёт просрочен, оплата не поступила',
      sections.overdue.length, '#fce8e6',
      'Период оказания услуги уже истёк, но деньги не пришли. Свяжитесь с плательщиком и выясните причину задержки.');
    writeTableHeader(['VIN', 'Модель / Гос.номер', 'Плательщик', 'Счёт просрочен на', 'Сумма за авто']);
    sections.overdue.sort(function(a,b){ return b.daysOverdue - a.daysOverdue; });
    sections.overdue.forEach(function(item) {
      writeRow([
        item.vin || '—',
        item.model + (item.reg ? ' / '+item.reg : ''),
        item.payer,
        item.daysOverdue + ' дн. (счёт ' + (item.invNum||'—') + (item.totalVins > 1 ? ', '+item.totalVins+' авто' : '') + ')',
        fmtMoney(item.amount)
      ], '#fff5f5');
    });
    writeEmpty();
  }

  // ────────────────────────────────────
  // СЕКЦИЯ 2: Заканчивается подключение
  // ────────────────────────────────────
  if (sections.expiringSoon.length > 0) {
    writeSectionHeader('🟡', 'ПОРА ВЫСТАВИТЬ СЧЁТ — подключение заканчивается',
      sections.expiringSoon.length, '#fef7e0',
      'Авто оплачено, но период подключения скоро истечёт. Выставите счёт заранее, чтобы не было перерыва в мониторинге.');
    writeTableHeader(['VIN', 'Модель / Гос.номер', 'Плательщик', 'Подключен до', 'Действие']);
    sections.expiringSoon.sort(function(a,b){ return (a.daysLeft||0) - (b.daysLeft||0); });
    sections.expiringSoon.forEach(function(item) {
      var action, bg;
      if (item.daysExpired) {
        action = '⚠ Истёк ' + item.daysExpired + ' дн. назад — выставить счёт!';
        bg = '#fff3cd';
      } else if (item.daysLeft === 0) {
        action = '⚠ Истекает СЕГОДНЯ — выставить счёт!';
        bg = '#fff3cd';
      } else {
        action = 'Осталось ' + item.daysLeft + ' дн. — выставить заранее';
        bg = '#fffdf0';
      }
      writeRow([
        item.vin || '—',
        item.model + (item.reg ? ' / '+item.reg : ''),
        item.payer,
        fmtDate(item.connTo),
        action
      ], bg);
    });
    writeEmpty();
  }

  // ────────────────────────────────────
  // СЕКЦИЯ 3: Ждём оплаты
  // ────────────────────────────────────
  if (sections.awaitPay.length > 0) {
    writeSectionHeader('⏳', 'ОЖИДАЕМ ОПЛАТУ — счёт выставлен, деньги ещё не пришли',
      sections.awaitPay.length, '#f0f4ff',
      'Счёт уже выставлен. Никаких действий не требуется, если срок ещё не истёк. Отметьте оплату когда деньги придут.');
    writeTableHeader(['VIN', 'Модель / Гос.номер', 'Плательщик', 'Счёт / Срок', 'Сумма за авто']);
    sections.awaitPay.sort(function(a,b){ return (a.daysLeft||0) - (b.daysLeft||0); });
    sections.awaitPay.forEach(function(item) {
      writeRow([
        item.vin || '—',
        item.model + (item.reg ? ' / '+item.reg : ''),
        item.payer,
        (item.invNum||'—') + (item.totalVins > 1 ? ' ('+item.totalVins+' авто)' : '') + ' · осталось ' + item.daysLeft + ' дн.',
        fmtMoney(item.amount)
      ], '#f8f9ff');
    });
    writeEmpty();
  }

  // ────────────────────────────────────
  // СЕКЦИЯ 4: Нет счетов
  // ────────────────────────────────────
  if (sections.noInvoice.length > 0) {
    writeSectionHeader('📝', 'НЕТ НИ ОДНОГО СЧЁТА — нужно выставить первый счёт',
      sections.noInvoice.length, '#f3f4f6',
      'Эти авто есть в базе, но счёт никогда не выставлялся. Возможно, новые подключения или пропущены при выставлении.');
    writeTableHeader(['VIN', 'Модель / Гос.номер', 'Плательщик', 'Действие', '']);
    sections.noInvoice.forEach(function(item) {
      writeRow([
        item.vin || '—',
        item.model + (item.reg ? ' / '+item.reg : ''),
        item.payer,
        'Выставить первый счёт', ''
      ], '#fafafa');
    });
    writeEmpty();
  }

  // ────────────────────────────────────
  // СЕКЦИЯ 5: Всё хорошо
  // ────────────────────────────────────
  if (sections.ok.length > 0) {
    writeSectionHeader('✅', 'ВСЁ В ПОРЯДКЕ — оплачено, подключение актуально',
      sections.ok.length, '#e6f4ea',
      'Эти авто оплачены и период подключения не заканчивается в ближайшие ' + NOTIFY_DAYS + ' дней.');
    writeTableHeader(['VIN', 'Модель / Гос.номер', 'Плательщик', 'Подключен до', '']);
    sections.ok.sort(function(a,b){ return new Date(a.connTo||0) - new Date(b.connTo||0); });
    sections.ok.forEach(function(item) {
      writeRow([item.vin||'—', item.model+(item.reg?' / '+item.reg:''), item.payer, fmtDate(item.connTo), ''], '#f6fef9');
    });
  }

  // Ширины колонок
  sh.setColumnWidth(1, 160);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 190);
  sh.setColumnWidth(4, 220);
  sh.setColumnWidth(5, 130);
  sh.setFrozenRows(2);
}


// ============================================================
// СТАТУС СЧЁТА (onEdit)
// ============================================================
function onEdit(e) {
  var sheet = e.range.getSheet();

  // Очищаем placeholder при вводе названия плательщика
  if (sheet.getName() === SHEET_PAYERS && e.range.getColumn() === 2) {
    var val = e.range.getValue();
    if (val && String(val).indexOf('введи название') === -1) {
      e.range.setFontColor('#000000').setFontStyle('normal');
      e.range.clearNote();
      // Обновляем dropdown и переназначаем авто
      var ss   = SpreadsheetApp.getActiveSpreadsheet();
      var shVeh = ss.getSheetByName(SHEET_VEHICLES);
      if (shVeh && shVeh.getLastRow() > 1) {
        applyPayerDropdown(shVeh, shVeh.getLastRow() - 1);
        autoAssignPayers_(shVeh, shVeh.getLastRow() - 1);
      }
    }
    return;
  }

  // Авторасчёт статуса в Счетах
  if (sheet.getName() !== SHEET_INVOICES) return;
  var row = e.range.getRow();
  if (row < 2) return;

  updateInvoiceStatus(sheet, row);

  // Если редактировали столбец I (Дата оплаты) — обновить "Подключен до"
  // onEdit срабатывает при любом изменении строки, поэтому проверяем столбец
  var col = e.range.getColumn();
  if (col === 9) { // I = 9 = Дата оплаты
    updateConnectedDates_();
  }
}

function updateInvoiceStatus(sh, row) {
  var today = new Date(); today.setHours(0,0,0,0);
  var vals     = sh.getRange(row,1,1,10).getValues()[0].map(safeVal);
  var periodTo = vals[5] ? new Date(vals[5]) : null;
  var paidDate = vals[8];
  if (!periodTo) return;

  var daysLeft = Math.round((periodTo - today) / 86400000);
  var status, bg;
  if (paidDate)               { status = '✅ Оплачен';                                     bg = '#e6f4ea'; }
  else if (periodTo < today)  { status = '🔴 Просрочен ' + Math.abs(daysLeft) + ' дн.';    bg = '#fce8e6'; }
  else if (daysLeft <= NOTIFY_DAYS) { status = '🟡 Истекает через ' + daysLeft + ' дн.';   bg = '#fef7e0'; }
  else                        { status = '⚪ Активен (' + daysLeft + ' дн.)';               bg = '#ffffff'; }

  sh.getRange(row,10).setValue(status);
  sh.getRange(row,1,1,10).setBackground(bg);
}

// ============================================================
// DROPDOWN ПЛАТЕЛЬЩИКОВ
// ============================================================
function applyPayerDropdown(sh, count) {
  if (count < 1) return;
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var shPay = ss.getSheetByName(SHEET_PAYERS);
  if (!shPay || shPay.getLastRow() < 2) return;

  var payerRange = shPay.getRange(2, 2, shPay.getLastRow()-1, 1);
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(payerRange, true)
    .setAllowInvalid(true)   // разрешаем «нестандартные» значения (не блокируем)
    .setHelpText('Выберите плательщика из списка')
    .build();
  sh.getRange(2, 8, count, 1).setDataValidation(rule);
}

function refreshPayerDropdowns() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_VEHICLES);
  if (!sh || sh.getLastRow() < 2) return;
  applyPayerDropdown(sh, sh.getLastRow()-1);
  SpreadsheetApp.getUi().alert('✅ Выпадающие списки обновлены!');
}

// ============================================================
// УТИЛИТЫ
// ============================================================
function getOrCreateSheet(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function formatHeader(sh, cols) {
  sh.getRange(1,1,1,cols)
    .setBackground('#1a73e8').setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(11);
  sh.setFrozenRows(1);
}

function formatDate(d) {
  if (!d) return '';
  return Utilities.formatDate(new Date(d), 'Europe/Moscow', 'dd.MM.yyyy');
}

function buildEmailBody(needInvoice, overdueList, today) {
  var html = '<div style="font-family:Arial,sans-serif;max-width:700px">';
  html += '<h2 style="background:#1a73e8;color:#fff;padding:12px 16px;border-radius:6px">🔔 Телематика: требуют внимания</h2>';
  html += '<p style="color:#666">Дата проверки: ' + formatDate(today) + '</p>';

  if (overdueList.length > 0) {
    html += '<h3 style="color:#d93025">🔴 Просроченные неоплаченные счета (' + overdueList.length + ')</h3>';
    html += '<table style="width:100%;border-collapse:collapse">';
    html += '<tr style="background:#fce8e6"><th style="padding:8px;border:1px solid #ddd">VIN</th><th style="padding:8px;border:1px solid #ddd">Плательщик</th><th style="padding:8px;border:1px solid #ddd">№ счета</th><th style="padding:8px;border:1px solid #ddd">Период по</th><th style="padding:8px;border:1px solid #ddd">Сумма</th></tr>';
    overdueList.forEach(function(item) {
      html += '<tr><td style="padding:8px;border:1px solid #ddd">' + item.vin + '</td><td style="padding:8px;border:1px solid #ddd">' + item.payer + '</td><td style="padding:8px;border:1px solid #ddd">' + (item.invoice||'—') + '</td><td style="padding:8px;border:1px solid #ddd;color:#d93025">' + item.periodTo + '</td><td style="padding:8px;border:1px solid #ddd">' + (item.amount ? item.amount.toLocaleString('ru') + ' ₽' : '—') + '</td></tr>';
    });
    html += '</table>';
  }

  if (needInvoice.length > 0) {
    html += '<h3 style="color:#f29900;margin-top:24px">🟡 Нужно выставить счёт (' + needInvoice.length + ')</h3>';
    html += '<table style="width:100%;border-collapse:collapse">';
    html += '<tr style="background:#fef7e0"><th style="padding:8px;border:1px solid #ddd">VIN</th><th style="padding:8px;border:1px solid #ddd">Плательщик</th><th style="padding:8px;border:1px solid #ddd">Период по</th><th style="padding:8px;border:1px solid #ddd">Причина</th></tr>';
    needInvoice.forEach(function(item) {
      html += '<tr><td style="padding:8px;border:1px solid #ddd">' + item.vin + '</td><td style="padding:8px;border:1px solid #ddd">' + item.payer + '</td><td style="padding:8px;border:1px solid #ddd">' + (item.periodTo||'—') + '</td><td style="padding:8px;border:1px solid #ddd;color:#f29900">' + item.reason + '</td></tr>';
    });
    html += '</table>';
  }

  html += '<p style="margin-top:24px;color:#888;font-size:12px">Автоматическое уведомление системы учёта телематики Igetis</p>';
  html += '</div>';
  return html;
}

// ============================================================
// SIDEBAR: СОЗДАНИЕ СЧЁТА
// ============================================================
function openInvoiceSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('InvoiceSidebar')
    .setTitle('Создание счёта')
    .setWidth(320);
  SpreadsheetApp.getUi().showSidebar(html);
}

// Возвращает список плательщиков с реальными названиями
function getPayersList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PAYERS);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 6).getValues()
    .filter(function(r){ return r[0] && r[1] && String(r[1]).indexOf('введи название')===-1; })
    .map(function(r){ return { id: String(r[0]), name: String(r[1]), contract: String(r[5]||'') }; });
}

// Возвращает все автомобили с нужными полями
function getAllVehicles() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_VEHICLES);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow()-1, 11).getValues()
    .filter(function(r){ return r[0]; })
    .map(function(r){
      return {
        id:           String(safeVal(r[0])),
        vin:          String(safeVal(r[1])),
        model:        String(safeVal(r[2])),
        regNum:       String(safeVal(r[3])),
        payerName:    String(safeVal(r[7])),
        contragentId: String(safeVal(r[10])),
        connectedTo:  safeVal(r[11]) || ''
      };
    });
}

// Генерирует следующий номер счёта вида ТЛМ-ГГГГ-NNN
function getNextInvoiceNumber() {
  var ss   = SpreadsheetApp.getActiveSpreadsheet();
  var sh   = ss.getSheetByName(SHEET_INVOICES);
  var year = new Date().getFullYear();
  var prefix = 'ТЛМ-' + year + '-';
  var max  = 0;
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues().forEach(function(r){
      var v = String(r[0]);
      if (v.indexOf(prefix) === 0) {
        var n = parseInt(v.replace(prefix,''), 10);
        if (!isNaN(n) && n > max) max = n;
      }
    });
  }
  return prefix + String(max+1).padStart(3, '0');
}

// Сохраняет массив счетов на лист Счета
function saveInvoices(invoices) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_INVOICES);
    if (!sh) return { ok: false, error: 'Лист «Счета» не найден. Запустите «Первичную настройку».' };

    invoices.forEach(function(inv) {
      sh.appendRow([
        inv.num,             // A: № счёта
        inv.payerId,         // B: ID плательщика
        inv.payerName,       // C: Название плательщика
        inv.vins,            // D: VIN список
        inv.dFrom,           // E: Период с
        inv.dTo,             // F: Период по
        inv.amount,          // G: Сумма
        inv.invDate,         // H: Дата выставления
        '',                  // I: Дата оплаты
        '⚪ Активен'         // J: Статус (пересчитается при onEdit)
      ]);
      // Сразу рассчитываем статус
      updateInvoiceStatus(sh, sh.getLastRow());
    });

    updateDashboard();
    updateConnectedDates_();
    return { ok: true, count: invoices.length };
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

// ============================================================
// ПАМЯТЬ ПЛАТЕЛЬЩИКА: запоминает цену и следующий период
// Хранится в UserProperties (привязано к пользователю Google)
// ============================================================
function getPayerMemory(payerId) {
  try {
    var props = PropertiesService.getUserProperties();
    var raw   = props.getProperty('payer_mem_' + payerId);
    return raw ? JSON.parse(raw) : null;
  } catch(e) { return null; }
}

function savePayerMemory(payerId, price, nextFrom, nextTo) {
  try {
    PropertiesService.getUserProperties().setProperty(
      'payer_mem_' + payerId,
      JSON.stringify({ price: price, nextFrom: nextFrom, nextTo: nextTo })
    );
  } catch(e) {}
}

// ============================================================
// ИМПОРТ SIDEBAR
// ============================================================
function openImportSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ImportSidebar')
    .setTitle('Импорт данных')
    .setWidth(380);
  SpreadsheetApp.getUi().showSidebar(html);
}

// URL для скачивания шаблона из Google Drive (файл должен быть там)
// Если файл ещё не загружен — возвращаем инструкцию
function getTemplateDownloadUrl() {
  // Ищем файл igetis_import_template.xlsx в Drive
  var files = DriveApp.getFilesByName('igetis_import_template.xlsx');
  if (files.hasNext()) {
    return files.next().getDownloadUrl();
  }
  return null; // null = пользователь получит подсказку
}

// Парсинг Excel из base64 (через Utilities)
function parseXlsxBase64(b64) {
  try {
    // Загружаем xlsx на Drive через UrlFetchApp (не требует Advanced Drive Service)
    var token = ScriptApp.getOAuthToken();
    var bytes = Utilities.base64Decode(b64);
    var blob  = Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'import.xlsx');

    // Шаг 1: загружаем как xlsx
    var uploadResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/upload/drive/v3/files?uploadType=media',
      {
        method: 'post',
        contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        payload: blob.getBytes(),
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );
    if (uploadResp.getResponseCode() !== 200) {
      throw new Error('Upload failed: ' + uploadResp.getContentText());
    }
    var fileId = JSON.parse(uploadResp.getContentText()).id;

    // Шаг 2: конвертируем в Google Sheets через copy
    var convertResp = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + fileId + '/copy',
      {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ name: '_tmp_igetis_import', mimeType: 'application/vnd.google-apps.spreadsheet' }),
        headers: { 'Authorization': 'Bearer ' + token },
        muteHttpExceptions: true
      }
    );
    if (convertResp.getResponseCode() !== 200) {
      // Удаляем загруженный файл
      UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId,
        { method: 'delete', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
      throw new Error('Convert failed: ' + convertResp.getContentText());
    }
    var sheetId = JSON.parse(convertResp.getContentText()).id;

    // Шаг 3: парсим
    var tmpSS = SpreadsheetApp.openById(sheetId);
    var result = parseImportSpreadsheet(tmpSS);

    // Шаг 4: удаляем оба временных файла
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + fileId,
      { method: 'delete', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });
    UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files/' + sheetId,
      { method: 'delete', headers: { 'Authorization': 'Bearer ' + token }, muteHttpExceptions: true });

    return result;
  } catch(e) {
    Logger.log('parseXlsxBase64 ERROR: ' + e.message);
    return { rows: [], type: 'vehicles', error: e.message };
  }
}

function parseImportSpreadsheet(ss) {
  var sheets = ss.getSheets();
  var type = 'vehicles';
  var targetSheet = null;

  for (var i = 0; i < sheets.length; i++) {
    var sname = sheets[i].getName().toLowerCase();
    if (sname.indexOf('плательщик') !== -1 || sname.indexOf('payer') !== -1) {
      targetSheet = sheets[i]; type = 'payers'; break;
    }
    if (sname.indexOf('авто') !== -1 || sname.indexOf('vehicle') !== -1) {
      targetSheet = sheets[i]; type = 'vehicles';
    }
  }
  if (!targetSheet) targetSheet = sheets[0];

  var lastRow = targetSheet.getLastRow();
  var lastCol = Math.max(targetSheet.getLastColumn(), 1);
  if (lastRow < 2) return { rows: [], type: type };

  // Читаем первые 10 строк для поиска заголовков
  var MARKERS = ['vin','марка','модель','model','название','name','плательщик','payer','гос','номер','телефон'];
  var scanRows = targetSheet.getRange(1, 1, Math.min(lastRow, 10), lastCol).getValues();
  var headerRowIdx = 0; // по умолчанию строка 1
  for (var ri = 0; ri < scanRows.length; ri++) {
    var rowStr = scanRows[ri].join(' ').toLowerCase();
    var hits = 0;
    for (var mi = 0; mi < MARKERS.length; mi++) {
      if (rowStr.indexOf(MARKERS[mi]) !== -1) hits++;
    }
    if (hits >= 2) { headerRowIdx = ri; break; }
  }

  // Заголовки из найденной строки
  var headerRow = scanRows[headerRowIdx].map(function(h) {
    return String(h).replace(/\s*\*/g, '').trim();
  });

  // Данные — читаем ВСЕ строки после заголовка отдельным запросом
  var dataStartRow = headerRowIdx + 2; // +1 за нумерацию с 1, +1 пропускаем заголовок
  var rows = [];
  if (dataStartRow <= lastRow) {
    var dataRange = targetSheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();
    dataRange.forEach(function(row) {
      if (row.every(function(c) { return !String(c).trim(); })) return; // пустая строка
      var first = String(row[0]);
      if (first.indexOf('⚠') !== -1) return;   // строка-предупреждение шаблона
      if (first.indexOf('X9F3') !== -1) return; // строка-пример шаблона
      var obj = {};
      headerRow.forEach(function(h, ci) {
        if (!h) return;
        var cell = row[ci];
        var val = '';
        if (cell instanceof Date && !isNaN(cell)) {
          val = Utilities.formatDate(cell, 'Europe/Moscow', 'dd.MM.yyyy');
        } else {
          val = String(safeVal(cell) || '').trim();
        }
        obj[h] = val;
      });
      if (Object.keys(obj).some(function(k) { return obj[k]; })) rows.push(obj);
    });
  }

  Logger.log('parseImport: sheet="' + targetSheet.getName() + '" type=' + type +
    ' headerRow=' + (headerRowIdx+1) + ' dataFrom=' + dataStartRow +
    ' lastRow=' + lastRow + ' headers=' + JSON.stringify(headerRow) +
    ' parsed=' + rows.length);

  return { rows: rows, type: type };
}

// Главная функция импорта
function importData(data) {
  try {
    if (data.type === 'payers') return importPayers_(data.rows);
    return importVehicles_(data.rows);
  } catch(e) {
    return { ok: false, error: e.message };
  }
}

function importPayers_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_PAYERS);
  if (!sh) return { ok: false, error: 'Лист «Плательщики» не найден' };

  var existing = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow()-1, 6).getValues().forEach(function(r) {
      var name = String(safeVal(r[1])).toLowerCase();
      if (name) existing[name] = true;
    });
  }

  var added = 0, skipped = 0;
  rows.forEach(function(r) {
    var name = (r['Название юрлица'] || r['name'] || '').trim();
    var contract = (r['Номер договора'] || r['contract'] || '').trim();
    if (!name) { skipped++; return; }
    // Проверяем дублирование по имени
    if (existing[name.toLowerCase()]) { skipped++; return; }

    var id = generateManualId_();
    sh.appendRow([
      id,
      name,
      r['Контактное лицо'] || r['contact'] || '',
      r['Телефон'] || r['phone'] || '',
      r['Email'] || r['email'] || '',
      contract
    ]);
    // Зелёный фон для ручных плательщиков
    sh.getRange(sh.getLastRow(), 1, 1, 6).setBackground('#e6f4ea');
    existing[name.toLowerCase()] = true;
    added++;
  });

  // Обновляем dropdown в автомобилях
  var shVeh = ss.getSheetByName(SHEET_VEHICLES);
  if (shVeh && shVeh.getLastRow() > 1) applyPayerDropdown(shVeh, shVeh.getLastRow()-1);

  return { ok: true, added: added, updated: 0, skipped: skipped };
}

function importVehicles_(rows) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_VEHICLES);
  if (!sh) return { ok: false, error: 'Лист «Автомобили» не найден' };

  // Карта существующих VIN
  var existingVins = {};
  if (sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues().forEach(function(r) {
      var vin = String(safeVal(r[1])).trim();
      var id  = String(safeVal(r[0])).trim();
      if (vin) existingVins[vin] = id;
    });
  }

  // Карта плательщиков: id -> name
  var shPay = ss.getSheetByName(SHEET_PAYERS);
  var payerById = {}, payerByName = {};
  if (shPay && shPay.getLastRow() > 1) {
    shPay.getRange(2, 1, shPay.getLastRow()-1, 2).getValues().forEach(function(r) {
      var id = String(safeVal(r[0])); var name = String(safeVal(r[1]));
      if (id) { payerById[id] = name; payerByName[name.toLowerCase()] = id; }
    });
  }

  var added = 0, skipped = 0;
  rows.forEach(function(r) {
    var vin   = (r['VIN'] || '').trim();
    var model = (r['Марка / Модель'] || r['Марка/Модель'] || '').trim();
    if (!vin || !model) { skipped++; return; }
    if (existingVins[vin]) { skipped++; return; } // уже есть — пропускаем

    // Резолвим плательщика — принимаем и ID и название
    var payerId = (r['ID Плательщика'] || '').trim();
    var payerName = '';
    if (payerId) {
      payerName = payerById[payerId] || payerId;
    } else {
      // Пробуем найти по имени
      var pname = (r['Плательщик'] || '').trim().toLowerCase();
      payerId = payerByName[pname] || '';
      payerName = pname ? (payerById[payerId] || pname) : '';
    }

    // Генерируем ID для ручных авто
    var vehicleId = 'MAN-VEH-' + Utilities.getUuid().substring(0, 8).toUpperCase();

    sh.appendRow([
      vehicleId,           // A: ID
      vin,                 // B: VIN
      model,               // C: Марка
      r['Гос. номер'] || r['Гос.номер'] || '',   // D
      r['Номер трекера'] || '',                   // E
      r['Дата монтажа']  || '',                   // F
      r['Дата подключения'] || '',                // G
      payerName,                                  // H: Плательщик (название)
      r['Эксплуатант'] || '',                     // I
      r['Комментарий'] || '',                     // J
      payerId,                                    // K: contragentId
      r['Подключен до'] || ''                     // L: Подключен до
    ]);
    // Светло-зелёный фон для ручных авто
    sh.getRange(sh.getLastRow(), 1, 1, 1).setBackground('#e6f4ea');
    existingVins[vin] = vehicleId;
    added++;
  });

  updateDashboard();
  return { ok: true, added: added, updated: 0, skipped: skipped };
}

// ============================================================
// РУЧНОЕ ДОБАВЛЕНИЕ (из Import Sidebar)
// ============================================================
function addVehicleManual(d) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_VEHICLES);
    if (!sh) return { ok: false, error: 'Лист «Автомобили» не найден' };

    var id = 'MAN-VEH-' + Utilities.getUuid().substring(0, 8).toUpperCase();
    sh.appendRow([
      id, d.vin, d.model, d.reg, d.tracker,
      d.mount, d.connect, d.payerName,
      d.operator, d.comment, d.payerId,
      d.connectedTo || ''
    ]);
    sh.getRange(sh.getLastRow(), 1, 1, 1).setBackground('#e6f4ea');
    applyPayerDropdown(sh, sh.getLastRow()-1);
    updateDashboard();
    return { ok: true, row: sh.getLastRow() };
  } catch(e) { return { ok: false, error: e.message }; }
}

function addPayerManual(d) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sh = ss.getSheetByName(SHEET_PAYERS);
    if (!sh) return { ok: false, error: 'Лист «Плательщики» не найден' };

    var id = generateManualId_();
    sh.appendRow([id, d.name, d.contact, '', d.email, d.contract]);
    var newRow = sh.getLastRow();
    // Телефон записываем отдельно как @TEXT чтобы +7... не воспринималось как формула
    if (d.phone) {
      sh.getRange(newRow, 4).setNumberFormat('@').setValue(d.phone);
    }
    sh.getRange(newRow, 1, 1, 6).setBackground('#e6f4ea');

    // Обновляем dropdown
    var shVeh = ss.getSheetByName(SHEET_VEHICLES);
    if (shVeh && shVeh.getLastRow() > 1) applyPayerDropdown(shVeh, shVeh.getLastRow()-1);

    return { ok: true, id: id };
  } catch(e) { return { ok: false, error: e.message }; }
}

// Генерирует MAN-YYYYMMDD-XXXX (короткий, читаемый)
function generateManualId_() {
  var d = new Date();
  var date = Utilities.formatDate(d, 'Europe/Moscow', 'yyyyMMdd');
  var rand = Utilities.getUuid().substring(0, 4).toUpperCase();
  return 'MAN-' + date + '-' + rand;
}

// ============================================================
// ОБНОВЛЕНИЕ "ПОДКЛЮЧЕН ДО" ИЗ ОПЛАЧЕННЫХ СЧЕТОВ
// Логика: берём все оплаченные счета, для каждого VIN
// находим максимальную дату "Период по" среди оплаченных → Подключен до
// ============================================================
function updateConnectedDates_() {
  try {
    var ss    = SpreadsheetApp.getActiveSpreadsheet();
    var shInv = ss.getSheetByName(SHEET_INVOICES);
    var shVeh = ss.getSheetByName(SHEET_VEHICLES);
    if (!shInv || !shVeh) return;
    if (shInv.getLastRow() < 2 || shVeh.getLastRow() < 2) return;

    // Читаем все счета
    var invRows = shInv.getRange(2, 1, shInv.getLastRow()-1, 10).getValues().map(function(r){return r.map(safeVal);});

    // Строим карту VIN -> максимальная дата "Период по" из ОПЛАЧЕННЫХ счетов
    var vinToDate = {}; // VIN -> Date
    invRows.forEach(function(inv) {
      var vinsStr = String(inv[3] || '');   // D: список VIN
      var periodTo = inv[5];                 // F: Период по
      var paidDate = inv[8];                 // I: Дата оплаты
      if (!paidDate || !periodTo || !vinsStr) return;

      var dt = new Date(periodTo);
      if (isNaN(dt.getTime())) return;

      // Счёт может покрывать несколько VIN (через запятую)
      vinsStr.split(/[,;]/).forEach(function(v) {
        var vin = v.trim();
        if (!vin) return;
        if (!vinToDate[vin] || dt > vinToDate[vin]) {
          vinToDate[vin] = dt;
        }
      });
    });

    if (Object.keys(vinToDate).length === 0) return;

    // Обновляем столбец L в Автомобилях
    var vehData = shVeh.getRange(2, 1, shVeh.getLastRow()-1, 12).getValues();
    var updated = 0;
    vehData.forEach(function(row, idx) {
      var vin = String(safeVal(row[1])).trim();
      if (!vin || !vinToDate[vin]) return;

      var newDate = vinToDate[vin];
      var currentVal = safeVal(row[11]);

      // Обновляем если: пусто, или новая дата позже текущей
      var shouldUpdate = !currentVal;
      if (!shouldUpdate && currentVal) {
        var current = new Date(currentVal);
        if (!isNaN(current.getTime()) && newDate > current) shouldUpdate = true;
      }

      if (shouldUpdate) {
        shVeh.getRange(idx + 2, 12).setValue(newDate);
        // Цветовая маркировка: красный если просрочено, зелёный если актуально
        var today = new Date(); today.setHours(0,0,0,0);
        var bg = newDate >= today ? '#e6f4ea' : '#fce8e6';
        shVeh.getRange(idx + 2, 12).setBackground(bg);
        updated++;
      }
    });

    Logger.log('updateConnectedDates_: обновлено ' + updated + ' авто');
  } catch(e) {
    Logger.log('updateConnectedDates_ error: ' + e.message);
  }
}

// Публичная обёртка для запуска из меню
function refreshConnectedDates() {
  updateConnectedDates_();
  SpreadsheetApp.getUi().alert('✅ Поле «Подключен до» обновлено из оплаченных счетов.');
}

// ============================================================
// МИГРАЦИЯ СХЕМЫ: добавляет новые столбцы к существующим листам
// Безопасно — не трогает данные, только добавляет заголовок если нет
// ============================================================
function migrateSchema() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  var changes = [];

  // --- Автомобили: проверяем наличие столбца L "Подключен до" ---
  var shVeh = ss.getSheetByName(SHEET_VEHICLES);
  if (shVeh) {
    var lastCol = shVeh.getLastColumn();
    var headers = shVeh.getRange(1, 1, 1, lastCol).getValues()[0];

    // Столбец K: ID Плательщика (contragentId)
    if (headers.indexOf('ID Плательщика (contragentId)') === -1 && lastCol < 11) {
      shVeh.getRange(1, 11).setValue('ID Плательщика (contragentId)');
      shVeh.getRange(1, 11).setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold');
      shVeh.getRange(1, 11).setFontColor('#888888'); // серый — служебный
      shVeh.setColumnWidth(11, 280);
      changes.push('Добавлен столбец K: ID Плательщика');
    }

    // Столбец L: Подключен до
    if (headers.indexOf('Подключен до') === -1) {
      var colL = Math.max(lastCol + 1, 12);
      var hdrCell = shVeh.getRange(1, colL);
      hdrCell.setValue('Подключен до');
      hdrCell.setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
      hdrCell.setNote('Дата до которой оплачено подключение. Заполняется из оплаченных счетов автоматически.');
      shVeh.setColumnWidth(colL, 130);
      changes.push('Добавлен столбец Подключен до (L)');
    }
  }

  // --- Плательщики: проверяем наличие столбца F "Номер договора" ---
  var shPay = ss.getSheetByName(SHEET_PAYERS);
  if (shPay) {
    var payHeaders = shPay.getRange(1, 1, 1, shPay.getLastColumn()).getValues()[0];
    if (payHeaders.indexOf('Номер договора') === -1) {
      var colF = Math.max(shPay.getLastColumn() + 1, 6);
      var fCell = shPay.getRange(1, colF);
      fCell.setValue('Номер договора');
      fCell.setBackground('#1a73e8').setFontColor('#ffffff').setFontWeight('bold').setFontSize(11);
      shPay.setColumnWidth(colF, 140);
      changes.push('Добавлен столбец Номер договора (F) в Плательщики');
    }
  }

  if (changes.length === 0) {
    ui.alert('✅ Схема уже актуальна — все столбцы на месте.');
  } else {
    ui.alert('✅ Схема обновлена:\n\n• ' + changes.join('\n• ') +
      '\n\nЗапусти «📅 Обновить «Подключен до»» чтобы заполнить из счетов.');
  }
}

// ============================================================
// ПОИСК ПО АВТОМОБИЛЯМ
// ============================================================
function openSearchSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SearchSidebar')
    .setTitle('Поиск по автомобилям')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function getAllVehiclesForSearch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_VEHICLES);
  if (!sh || sh.getLastRow() < 2) return [];

  var data = sh.getRange(2, 1, sh.getLastRow()-1, 12).getValues();
  var result = [];
  data.forEach(function(row, idx) {
    var id  = String(safeVal(row[0])).trim();
    var vin = String(safeVal(row[1])).trim();
    if (!id && !vin) return;
    result.push({
      row:         idx + 2,
      id:          id,
      vin:         vin,
      model:       String(safeVal(row[2])).trim(),
      reg:         String(safeVal(row[3])).trim(),
      tracker:     String(safeVal(row[4])).trim(),
      payer:       String(safeVal(row[7])).trim(),
      connectedTo: row[11] ? Utilities.formatDate(new Date(row[11]), 'Europe/Moscow', 'yyyy-MM-dd') : ''
    });
  });
  return result;
}

// Прокрутить лист Автомобили к нужной строке и подсветить её
function navigateToRow(row) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_VEHICLES);
  if (!sh) return;
  ss.setActiveSheet(sh);
  var range = sh.getRange(row, 1, 1, 12);
  sh.setActiveRange(range);
  // Кратковременная подсветка строки
  var orig = range.getBackgrounds()[0];
  range.setBackground('#fff3cd');
  Utilities.sleep(600);
  // Восстанавливаем оригинальные цвета
  orig.forEach(function(bg, i) {
    sh.getRange(row, i+1).setBackground(bg || null);
  });
}

// ============================================================
// ДИАГНОСТИКА: найти устройство по VIN в сыром ответе API
function debugVIN() {
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt('Диагностика VIN', 'Введи VIN для поиска в ответе API:', ui.ButtonSet.OK_CANCEL);
  if (result.getSelectedButton() !== ui.Button.OK) return;
  var searchVin = result.getResponseText().trim().toUpperCase();
  if (!searchVin) return;

  var options = { method: 'get', headers: { 'Authorization': 'Bearer ' + IGETIS_TOKEN }, muteHttpExceptions: true };
  var resp = UrlFetchApp.fetch('https://api.lk.igetis.pro/v1/devices/my?IsShowDisabled=true', options);
  if (resp.getResponseCode() !== 200) { ui.alert('Ошибка API: ' + resp.getResponseCode()); return; }

  var devices = JSON.parse(resp.getContentText()).results || [];
  var totalDevices = devices.length;
  var vinFields = ['vinNumber','vin','VIN','vinCode','vin_number'];
  var found = null;
  var matchField = '';

  devices.forEach(function(d) {
    if (found) return;
    if (JSON.stringify(d).toUpperCase().indexOf(searchVin) === -1) return;
    found = d;
    vinFields.forEach(function(f) {
      if (d[f] && String(d[f]).toUpperCase().indexOf(searchVin) !== -1) matchField = f;
    });
    if (!matchField) matchField = '(в JSON, не в стандартных полях)';
  });

  Logger.log('debugVIN — устройств в API: ' + totalDevices);
  if (found) Logger.log('Найдено: ' + JSON.stringify(found, null, 2));

  if (!found) {
    ui.alert('VIN ' + searchVin + ' не найден в API.' +
      '\n\nВсего устройств: ' + totalDevices +
      '\n\nВозможные причины:' +
      '\n1. VIN не внесён в ЛК Igetis' +
      '\n2. Устройство на другом контрагенте' +
      '\n3. Токен без прав на это устройство' +
      '\n\nРешение: добавь вручную через Импорт'
    );
    return;
  }

  var info = 'НАЙДЕНО. Поле VIN: ' + matchField;
  info += '\n\ndeviceId:     ' + (found.deviceId || 'нет');
  info += '\nvinNumber:    ' + (found.vinNumber || '!!! ПУСТО !!!');
  info += '\ncarModel:     ' + (found.carModel || 'нет');
  info += '\ncarNumber:    ' + (found.carNumber || 'нет');
  info += '\ndeviceNumber: ' + (found.deviceNumber || 'нет');
  info += '\ncontragentId: ' + (found.contragentId || 'нет');
  info += '\nisEnabled:    ' + found.isEnabled;

  if (matchField !== 'vinNumber') {
    info += '\n\n!!! ПРИЧИНА: VIN хранится в поле "' + matchField + '", а код читает "vinNumber" !!!';
    info += '\nОбратись к разработчику: нужно читать d.' + matchField;
  } else if (!found.vinNumber) {
    info += '\n\n!!! ПРИЧИНА: поле vinNumber ПУСТО — авто пропускается при синхронизации !!!';
    info += '\nЗаполни VIN в личном кабинете Igetis или добавь вручную.';
  } else {
    info += '\n\n✅ vinNumber заполнен. После синхронизации авто должно появиться.';
  }

  ui.alert(info);
}


// ============================================================
// АУДИТ БАЗЫ ДАННЫХ
// ============================================================
function runAudit() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var shVeh = ss.getSheetByName(SHEET_VEHICLES);
  var shPay = ss.getSheetByName(SHEET_PAYERS);
  var shInv = ss.getSheetByName(SHEET_INVOICES);

  // Полный URL таблицы для HYPERLINK-навигации
  var ssUrl   = ss.getUrl().replace(/\/edit.*$/, '/edit');
  var gidVeh  = shVeh ? shVeh.getSheetId() : 0;
  var gidPay  = shPay ? shPay.getSheetId() : 0;
  var gidInv  = shInv ? shInv.getSheetId() : 0;

  var issues = [];
  var stats  = {};
  var today  = new Date(); today.setHours(0,0,0,0);

  // ===== АВТОМОБИЛИ =====
  var vehicles = [];
  if (shVeh && shVeh.getLastRow() > 1) {
    vehicles = shVeh.getRange(2, 1, shVeh.getLastRow()-1, 12).getValues()
      .map(function(r,i){ return {
        row:i+2, id:safeVal(r[0]), vin:String(safeVal(r[1])).trim(),
        model:String(safeVal(r[2])).trim(), reg:String(safeVal(r[3])).trim(),
        tracker:String(safeVal(r[4])).trim(), payer:String(safeVal(r[7])).trim(),
        payerId:String(safeVal(r[10])).trim(), connTo:safeVal(r[11]),
        gid: gidVeh
      };})
      .filter(function(v){ return v.id || v.vin; });
  }
  stats.totalVehicles = vehicles.length;

  var vinCount = {};
  vehicles.forEach(function(v){ if(v.vin) vinCount[v.vin]=(vinCount[v.vin]||0)+1; });
  Object.keys(vinCount).forEach(function(vin){
    if(vinCount[vin]>1) issues.push({sev:'🔴',cat:'Авто',msg:'Дубль VIN: '+vin+' ('+vinCount[vin]+' раза)',sheet:SHEET_VEHICLES,gid:gidVeh});
  });

  var trackerCount = {};
  vehicles.forEach(function(v){ if(v.tracker) trackerCount[v.tracker]=(trackerCount[v.tracker]||0)+1; });
  Object.keys(trackerCount).forEach(function(t){
    if(trackerCount[t]>1) issues.push({sev:'🔴',cat:'Авто',msg:'Дубль трекера: '+t+' ('+trackerCount[t]+' авто)',sheet:SHEET_VEHICLES,gid:gidVeh});
  });

  var noVin = vehicles.filter(function(v){ return !v.vin; });
  stats.noVin = noVin.length;
  noVin.forEach(function(v){ issues.push({sev:'🟡',cat:'Авто',msg:'Нет VIN: '+v.model,sheet:SHEET_VEHICLES,row:v.row,gid:gidVeh}); });

  var noPayer = vehicles.filter(function(v){ return !v.payer; });
  stats.noPayer = noPayer.length;
  noPayer.forEach(function(v){ issues.push({sev:'🟡',cat:'Авто',msg:'Нет плательщика: '+(v.vin||v.model),sheet:SHEET_VEHICLES,row:v.row,gid:gidVeh}); });

  stats.noTracker = vehicles.filter(function(v){ return !v.tracker; }).length;

  var badVin = vehicles.filter(function(v){ return v.vin && !/^[A-HJ-NPR-Z0-9]{17}$/i.test(v.vin); });
  badVin.forEach(function(v){ issues.push({sev:'🟡',cat:'Авто',msg:'Некорректный VIN (не 17 симв.): '+v.vin,sheet:SHEET_VEHICLES,row:v.row,gid:gidVeh}); });

  var expired = vehicles.filter(function(v){
    if(!v.connTo) return false;
    var d=new Date(v.connTo); return !isNaN(d)&&d<today;
  });
  stats.expired = expired.length;
  expired.forEach(function(v){
    var days=Math.round((today-new Date(v.connTo))/86400000);
    issues.push({sev:'🔴',cat:'Оплата',msg:'Истёк '+days+' дн. назад: '+(v.vin||v.model)+' / '+v.payer,sheet:SHEET_VEHICLES,row:v.row,gid:gidVeh});
  });
  stats.noConnTo = vehicles.filter(function(v){ return !v.connTo; }).length;

  // ===== ПЛАТЕЛЬЩИКИ =====
  var payers = [];
  if (shPay && shPay.getLastRow() > 1) {
    payers = shPay.getRange(2,1,shPay.getLastRow()-1,6).getValues()
      .map(function(r,i){ return {row:i+2,id:String(safeVal(r[0])).trim(),name:String(safeVal(r[1])).trim(),contract:String(safeVal(r[5]||'')).trim(),gid:gidPay}; })
      .filter(function(p){ return p.id; });
  }
  stats.totalPayers = payers.length;

  var payerIds = {};
  vehicles.forEach(function(v){
    if (v.payerId) payerIds[v.payerId.trim().toLowerCase()] = true;
    if (v.payer)   payerIds[v.payer.trim().toLowerCase()]   = true;
  });
  payers.filter(function(p){
    return !payerIds[p.id.trim().toLowerCase()] && !payerIds[p.name.trim().toLowerCase()];
  })
    .forEach(function(p){ issues.push({sev:'⚪',cat:'Плательщик',msg:'Нет авто у плательщика: '+p.name,sheet:SHEET_PAYERS,row:p.row,gid:gidPay}); });
  payers.filter(function(p){ return p.id.indexOf('MAN-')===0&&!p.contract; })
    .forEach(function(p){ issues.push({sev:'🟡',cat:'Плательщик',msg:'Нет договора: '+p.name,sheet:SHEET_PAYERS,row:p.row,gid:gidPay}); });

  // ===== СЧЕТА =====
  var invoices = [];
  if (shInv && shInv.getLastRow() > 1) {
    invoices = shInv.getRange(2,1,shInv.getLastRow()-1,10).getValues()
      .map(function(r,i){ return {row:i+2,num:String(safeVal(r[0])),payerName:String(safeVal(r[2])),
        vins:String(safeVal(r[3])),dFrom:safeVal(r[4]),dTo:safeVal(r[5]),
        amount:safeVal(r[6]),invDate:safeVal(r[7]),paidDate:safeVal(r[8]),gid:gidInv}; })
      .filter(function(inv){ return inv.num; });
  }
  stats.totalInvoices = invoices.length;
  stats.paid   = invoices.filter(function(i){ return i.paidDate; }).length;
  stats.unpaid = invoices.filter(function(i){ return !i.paidDate; }).length;

  var overdueInv = invoices.filter(function(inv){
    if(inv.paidDate) return false;
    var dt=new Date(inv.dTo); return !isNaN(dt)&&dt<today;
  });
  stats.overdueInvoices = overdueInv.length;
  overdueInv.forEach(function(inv){ issues.push({sev:'🔴',cat:'Счёт',msg:'Просрочен и не оплачен: '+inv.num+' / '+inv.payerName,sheet:SHEET_INVOICES,row:inv.row,gid:gidInv}); });

  var vinSet = {};
  vehicles.forEach(function(v){ if(v.vin) vinSet[v.vin]=true; });
  invoices.forEach(function(inv){
    (inv.vins||'').split(/[,;]/).forEach(function(v){
      var vin=v.trim();
      if(vin&&!vinSet[vin]) issues.push({sev:'🟡',cat:'Счёт',msg:'VIN в счёте не найден в базе: '+vin+' (счёт '+inv.num+')',sheet:SHEET_INVOICES,row:inv.row,gid:gidInv});
    });
  });

  // ===== ЗАПИСЫВАЕМ ОТЧЁТ (6 колонок) =====
  var shName = '📋 Аудит';
  var shAudit = ss.getSheetByName(shName) || ss.insertSheet(shName);
  shAudit.clearContents(); shAudit.clearFormats();

  var now = Utilities.formatDate(new Date(),'Europe/Moscow','dd.MM.yyyy HH:mm');
  var red    = issues.filter(function(i){return i.sev==='🔴';}).length;
  var yellow = issues.filter(function(i){return i.sev==='🟡';}).length;
  var white  = issues.filter(function(i){return i.sev==='⚪';}).length;

  // Статичные строки (без ссылок) — записываем через setValues
  var staticRows = [
    ['📋 ОТЧЁТ АУДИТА — Igetis Telematics','','','','',now],          // 1
    ['','','','','',''],                                                  // 2
    ['СВОДКА','','','','',''],                                           // 3
    ['Автомобилей в базе:',stats.totalVehicles,'Без VIN:',stats.noVin,'',''],
    ['Без плательщика:',stats.noPayer,'Без трекера:',stats.noTracker,'',''],
    ['Подключение истекло:',stats.expired,'Без даты подключения:',stats.noConnTo,'',''],
    ['Плательщиков:',stats.totalPayers,'','','',''],
    ['Счетов всего:',stats.totalInvoices,'Оплачено:',stats.paid,'Просрочено: '+stats.overdueInvoices,''],
    ['','','','','',''],                                                  // 9
    ['ПРОБЛЕМЫ:  🔴 Критичных: '+red+'   🟡 Предупреждений: '+yellow+'   ⚪ Информационных: '+white,'','','','',''], // 10
    ['','','','','',''],                                                  // 11
  ];

  // Сортируем: 🔴 → 🟡 → ⚪; внутри «Оплата» — по убыванию дней просрочки
  issues.sort(function(a, b) {
    var order = {'🔴': 0, '🟡': 1, '⚪': 2};
    var sa = order[a.sev] !== undefined ? order[a.sev] : 9;
    var sb = order[b.sev] !== undefined ? order[b.sev] : 9;
    if (sa !== sb) return sa - sb;
    if (a.cat === 'Оплата' && b.cat === 'Оплата') {
      var da = parseInt((a.msg.match(/(\d+)\s*дн/) || [0,0])[1]) || 0;
      var db = parseInt((b.msg.match(/(\d+)\s*дн/) || [0,0])[1]) || 0;
      return db - da;
    }
    return 0;
  });

  if (issues.length === 0) {
    staticRows.push(['✅ Проблем не обнаружено!','','','','','']);
    shAudit.getRange(1,1,staticRows.length,6).setValues(staticRows);
  } else {
    staticRows.push(['Серьёзность','Категория','Описание','Лист','Строка','Перейти']); // 12 = headerRow
    shAudit.getRange(1,1,staticRows.length,6).setValues(staticRows);

    var headerRow = staticRows.length; // строка 12

    // Записываем строки проблем — данные + формулы ссылок отдельно
    issues.forEach(function(iss, idx) {
      var dataRow = headerRow + 1 + idx;
      var rowNum  = iss.row || '';
      shAudit.getRange(dataRow, 1, 1, 5).setValues([[iss.sev, iss.cat, iss.msg, iss.sheet||'', rowNum]]);

      // Колонка 6: кликабельная ссылка если есть конкретная строка
      if (iss.row && iss.gid !== undefined) {
        var cellAddr = 'A' + iss.row;
        var url = ssUrl + '#gid=' + iss.gid + '&range=' + cellAddr;
        var formula = '=HYPERLINK("' + url + '";">> ' + iss.row + '")';
        shAudit.getRange(dataRow, 6).setFormula(formula)
          .setFontColor('#1a73e8').setFontWeight('bold');
      } else {
        shAudit.getRange(dataRow, 6).setValue('');
      }
    });

    // Форматирование строк по цвету
    issues.forEach(function(iss, idx) {
      var dataRow = headerRow + 1 + idx;
      var bg = iss.sev==='🔴' ? '#fce8e6' : iss.sev==='🟡' ? '#fff3cd' : '#f9f9f9';
      shAudit.getRange(dataRow, 1, 1, 6).setBackground(bg);
    });

    // Заголовок таблицы
    shAudit.getRange(headerRow,1,1,6).setBackground('#f1f3f4').setFontWeight('bold');
  }

  // Форматирование шапки и сводки
  shAudit.getRange(1,1,1,6).merge().setBackground('#1a73e8').setFontColor('#fff').setFontWeight('bold').setFontSize(13).setVerticalAlignment('middle');
  shAudit.setRowHeight(1, 32);
  shAudit.getRange(3,1,1,6).setBackground('#e8f0fe').setFontWeight('bold');
  shAudit.getRange(10,1,1,6).setFontWeight('bold');

  // Ширины
  shAudit.setColumnWidth(1, 22);   // Серьёзность
  shAudit.setColumnWidth(2, 110);  // Категория
  shAudit.setColumnWidth(3, 370);  // Описание
  shAudit.setColumnWidth(4, 130);  // Лист
  shAudit.setColumnWidth(5, 60);   // Строка
  shAudit.setColumnWidth(6, 90);   // Перейти
  shAudit.setFrozenRows(1);

  ss.setActiveSheet(shAudit);
  SpreadsheetApp.getUi().alert(
    'Аудит завершён\n\n' +
    '🔴 Критичных: ' + red + '\n' +
    '🟡 Предупреждений: ' + yellow + '\n' +
    '⚪ Информационных: ' + white + '\n\n' +
    'Подробности — на листе «📋 Аудит»\n' +
    'Клик по «→ строка N» откроет нужную запись.'
  );
}



// ============================================================
// ВОССТАНОВЛЕНИЕ: найти и вернуть пропавшие ручные авто
// Читает лист целиком включая "пустые" строки с данными
// ============================================================
function recoverLostVehicles() {
  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var shV = ss.getSheetByName(SHEET_VEHICLES);
  if (!shV) return;

  var ui = SpreadsheetApp.getUi();

  // Читаем ВСЕ строки включая те что кажутся пустыми
  var maxRow = Math.min(shV.getMaxRows(), 10000);
  var allData = shV.getRange(2, 1, maxRow - 1, 12).getValues();
  
  // Строки где есть хоть что-то, но не видны в обычном диапазоне
  var lastVisible = shV.getLastRow();
  var lost = [];
  
  allData.forEach(function(row, idx) {
    var realRow = idx + 2;
    if (realRow <= lastVisible) return; // видимая строка — не трогаем
    var id  = String(row[0] || '').trim();
    var vin = String(row[1] || '').trim();
    if (!id && !vin) return; // реально пустая
    lost.push({ row: realRow, data: row });
  });

  if (lost.length === 0) {
    ui.alert('Пропавших авто не найдено.\n\n' +
      'Если авто исчезло — введи его заново через\n' +
      'меню Импорт, вкладка Добавить авто.');
    return;
  }

  // Добавляем найденные строки в конец видимых данных
  lost.forEach(function(item) {
    shV.appendRow(item.data);
    shV.getRange(shV.getLastRow(), 1, 1, 12).setBackground('#e6f4ea');
  });

  ui.alert('Восстановлено авто: ' + lost.length + '\n\n' +
    'Добавлены в конец списка с зелёным фоном.\n' +
    'Запусти синхронизацию с API чтобы упорядочить список.');
}

// ============================================================
// ДИАГНОСТИКА ИМПОРТА — запускать вручную из редактора Apps Script
// Перед запуском: загрузи файл на Google Drive и вставь его ID ниже
// ============================================================
function debugImportFile() {
  var ui = SpreadsheetApp.getUi();

  // Ищем последний временный файл импорта если он не был удалён
  var files = DriveApp.getFilesByName('_tmp_igetis_import');
  var tmpSS = null;

  if (files.hasNext()) {
    var f = files.next();
    tmpSS = SpreadsheetApp.openById(f.getId());
    Logger.log('Найден временный файл: ' + f.getId());
  } else {
    ui.alert('Временный файл не найден.\n\n' +
      'Попробуй так:\n' +
      '1. Загрузи файл через сайдбар (получишь ошибку)\n' +
      '2. НЕ закрывай редактор\n' +
      '3. Сразу запусти debugImportFile — временный файл ещё может быть на диске\n\n' +
      'Или: загрузи свой xlsx на Google Drive, открой его как Google Sheets\n' +
      'и вставь ID таблицы в код функции.');
    return;
  }

  var sheets = tmpSS.getSheets();
  var info = 'Листов в файле: ' + sheets.length + '\n\n';

  sheets.forEach(function(sh, idx) {
    info += 'Лист ' + (idx+1) + ': ' + sh.getName() + '\n';
    info += '  Строк: ' + sh.getLastRow() + ', Колонок: ' + sh.getLastColumn() + '\n';
    if (sh.getLastRow() > 0 && sh.getLastColumn() > 0) {
      var sample = sh.getRange(1, 1, Math.min(sh.getLastRow(), 5), Math.min(sh.getLastColumn(), 5)).getValues();
      sample.forEach(function(row, ri) {
        info += '  Строка ' + (ri+1) + ': ' + JSON.stringify(row.map(function(c){ return c instanceof Date ? c.toLocaleDateString() : String(c).substring(0,20); })) + '\n';
      });
    }
    info += '\n';
  });

  // Запускаем parseImportSpreadsheet и смотрим результат
  var result = parseImportSpreadsheet(tmpSS);
  info += 'Результат парсинга:\n';
  info += '  type: ' + result.type + '\n';
  info += '  rows: ' + result.rows.length + '\n';
  if (result.rows.length > 0) {
    info += '  Первая строка: ' + JSON.stringify(result.rows[0]) + '\n';
  }

  Logger.log(info);
  ui.alert(info.substring(0, 1500));
}
