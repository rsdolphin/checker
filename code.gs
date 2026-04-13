/**
 * Google Drive Security Scanner
 * Сканирует Google Drive на предмет файлов с публичным доступом
 */

// Константы для настройки
var CONFIG = {
  SHEET_NAME: 'Результаты сканирования',
  BATCH_SIZE: 50,
  MAX_EXECUTION_TIME: 300000, // 5 минут (оставляем запас до таймаута в 6 минут)
  PROPERTY_KEY: 'scanProgress',
  LOG_SHEET_NAME: 'Логи'
};

/**
 * Создает кастомное меню при открытии таблицы
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔒 Проверка безопасности')
    .addItem('🔍 Сканировать Drive', 'startSecurityScan')
    .addItem('📊 Очистить результаты', 'clearResults')
    .addItem('🔄 Сбросить прогресс', 'resetProgress')
    .addItem('📋 Показать статус', 'showStatus')
    .addToUi();
}

/**
 * Запускает сканирование безопасности
 */
function startSecurityScan() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Начать сканирование?',
    'Скрипт просканирует весь ваш Google Drive на предмет публично доступных файлов. ' +
    'Это может занять несколько минут.',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response == ui.Button.OK) {
    try {
      initializeSheets();
      logMessage('Запуск сканирования...');
      continueScan();
    } catch (e) {
      ui.alert('Ошибка: ' + e.message);
      logMessage('Ошибка: ' + e.message);
    }
  }
}

/**
 * Инициализирует таблицы для результатов и логов
 */
function initializeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Создаем/очищаем лист результатов
  var resultsSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!resultsSheet) {
    resultsSheet = ss.insertSheet(CONFIG.SHEET_NAME);
  } else {
    resultsSheet.clear();
  }
  
  // Заголовки таблицы
  var headers = [
    'Имя файла',
    'Тип файла',
    'Ссылка',
    'Владелец',
    'Тип доступа',
    'Уровень доступа',
    'Дата обнаружения'
  ];
  resultsSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  resultsSheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  resultsSheet.setFrozenRows(1);
  
  // Создаем лист логов
  var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
  } else {
    logSheet.clear();
  }
  
  var logHeaders = ['Время', 'Сообщение'];
  logSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
  logSheet.getRange(1, 1, 1, logHeaders.length).setFontWeight('bold');
  logSheet.setFrozenRows(1);
  
  // Сбрасываем прогресс
  PropertiesService.getScriptProperties().deleteProperty(CONFIG.PROPERTY_KEY);
}

/**
 * Продолжает сканирование с учетом прогресса
 */
function continueScan() {
  var startTime = new Date().getTime();
  var properties = PropertiesService.getScriptProperties();
  var progress = properties.getProperty(CONFIG.PROPERTY_KEY);
  
  var scanState;
  if (progress) {
    scanState = JSON.parse(progress);
    logMessage('Продолжение сканирования... Обработано: ' + scanState.processedCount + ', Найдено: ' + scanState.foundCount);
  } else {
    scanState = {
      processedCount: 0,
      foundCount: 0,
      folderQueue: ['root'],
      processedFolders: {},
      continuationToken: null
    };
    logMessage('Начало нового сканирования...');
  }
  
  var results = [];
  var filesProcessed = 0;
  
  try {
    // Обрабатываем файлы порциями
    while (scanState.folderQueue.length > 0) {
      var currentTime = new Date().getTime();
      
      // Проверяем таймаут
      if (currentTime - startTime > CONFIG.MAX_EXECUTION_TIME) {
        logMessage('Достигнут лимит времени выполнения. Сохранение прогресса...');
        break;
      }
      
      var folderId = scanState.folderQueue.shift();
      
      // Пропускаем уже обработанные папки
      if (scanState.processedFolders[folderId]) {
        continue;
      }
      
      scanState.processedFolders[folderId] = true;
      
      try {
        var folder = folderId === 'root' ? DriveApp.getRootFolder() : DriveApp.getFolderById(folderId);
        
        // Сканируем файлы в текущей папке
        var fileResults = scanFilesInFolder(folder);
        results = results.concat(fileResults.files);
        scanState.processedCount += fileResults.processed;
        scanState.foundCount += fileResults.found;
        filesProcessed += fileResults.processed;
        
        // Добавляем подпапки в очередь
        var subfolders = folder.getFolders();
        while (subfolders.hasNext()) {
          var subfolder = subfolders.next();
          var subfolderId = subfolder.getId();
          if (!scanState.processedFolders[subfolderId]) {
            scanState.folderQueue.push(subfolderId);
          }
        }
        
      } catch (e) {
        logMessage('Ошибка при обработке папки ' + folderId + ': ' + e.message);
      }
      
      // Периодически сохраняем результаты
      if (results.length >= CONFIG.BATCH_SIZE) {
        saveResults(results);
        results = [];
      }
    }
    
    // Сохраняем оставшиеся результаты
    if (results.length > 0) {
      saveResults(results);
    }
    
    // Проверяем завершение
    if (scanState.folderQueue.length === 0) {
      // Сканирование завершено
      properties.deleteProperty(CONFIG.PROPERTY_KEY);
      logMessage('✅ Сканирование завершено! Обработано файлов: ' + scanState.processedCount + ', Найдено проблем: ' + scanState.foundCount);
      
      SpreadsheetApp.getUi().alert(
        'Сканирование завершено',
        'Обработано файлов: ' + scanState.processedCount + '\n' +
        'Найдено файлов с публичным доступом: ' + scanState.foundCount,
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    } else {
      // Сохраняем прогресс и создаем триггер для продолжения
      properties.setProperty(CONFIG.PROPERTY_KEY, JSON.stringify(scanState));
      logMessage('Сохранен прогресс. Осталось папок: ' + scanState.folderQueue.length);
      
      // Создаем триггер для продолжения через несколько секунд
      ScriptApp.newTrigger('continueScan')
        .timeBased()
        .after(1000)
        .create();
    }
    
  } catch (e) {
    logMessage('Критическая ошибка: ' + e.message);
    properties.setProperty(CONFIG.PROPERTY_KEY, JSON.stringify(scanState));
    throw e;
  }
}

/**
 * Сканирует файлы в конкретной папке
 */
function scanFilesInFolder(folder) {
  var files = folder.getFiles();
  var results = [];
  var processed = 0;
  var found = 0;
  
  while (files.hasNext()) {
    try {
      var file = files.next();
      processed++;
      
      var sharingAccess = file.getSharingAccess();
      var sharingPermission = file.getSharingPermission();
      
      // Проверяем публичный доступ
      if (isPublicAccess(sharingAccess, sharingPermission)) {
        found++;
        
        var accessType = getAccessTypeDescription(sharingAccess, sharingPermission);
        var owner = 'Неизвестен';
        
        try {
          owner = file.getOwner() ? file.getOwner().getEmail() : 'Неизвестен';
        } catch (e) {
          // Если нет прав на получение владельца
          owner = 'Нет доступа';
        }
        
        results.push({
          name: file.getName(),
          mimeType: getMimeTypeDescription(file.getMimeType()),
          url: file.getUrl(),
          owner: owner,
          sharingAccess: sharingAccess.toString(),
          sharingPermission: sharingPermission.toString(),
          date: new Date()
        });
      }
    } catch (e) {
      logMessage('Ошибка при обработке файла: ' + e.message);
    }
  }
  
  return {
    files: results,
    processed: processed,
    found: found
  };
}

/**
 * Проверяет, является ли доступ публичным
 */
function isPublicAccess(sharingAccess, sharingPermission) {
  // ANYONE = Anyone with the link
  // ANYONE_WITH_LINK = Anyone with the link
  // DOMAIN = Anyone in the domain
  // DOMAIN_WITH_LINK = Anyone in the domain with the link
  
  return (
    sharingAccess === DriveApp.Access.ANYONE ||
    sharingAccess === DriveApp.Access.ANYONE_WITH_LINK ||
    sharingAccess === DriveApp.Access.DOMAIN ||
    sharingAccess === DriveApp.Access.DOMAIN_WITH_LINK
  );
}

/**
 * Получает описание типа доступа
 */
function getAccessTypeDescription(sharingAccess, sharingPermission) {
  var accessMap = {
    'ANYONE': 'Публично в интернете',
    'ANYONE_WITH_LINK': 'Доступно по ссылке (любой)',
    'DOMAIN': 'Доступно в домене',
    'DOMAIN_WITH_LINK': 'Доступно в домене по ссылке',
    'PRIVATE': 'Приватно'
  };
  
  var permissionMap = {
    'VIEW': '(просмотр)',
    'EDIT': '(редактирование)',
    'COMMENT': '(комментирование)',
    'OWNER': '(владелец)',
    'ORGANIZER': '(организатор)',
    'FILE_ORGANIZER': '(файловый организатор)',
    'NONE': '(нет прав)'
  };
  
  var access = accessMap[sharingAccess.toString()] || sharingAccess.toString();
  var permission = permissionMap[sharingPermission.toString()] || sharingPermission.toString();
  
  return access + ' ' + permission;
}

/**
 * Получает читаемое описание MIME типа
 */
function getMimeTypeDescription(mimeType) {
  var typeMap = {
    'application/vnd.google-apps.document': 'Google Документ',
    'application/vnd.google-apps.spreadsheet': 'Google Таблица',
    'application/vnd.google-apps.presentation': 'Google Презентация',
    'application/vnd.google-apps.form': 'Google Форма',
    'application/vnd.google-apps.drawing': 'Google Рисунок',
    'application/vnd.google-apps.folder': 'Папка',
    'application/pdf': 'PDF',
    'image/jpeg': 'JPEG изображение',
    'image/png': 'PNG изображение',
    'image/gif': 'GIF изображение',
    'video/mp4': 'MP4 видео',
    'application/zip': 'ZIP архив',
    'text/plain': 'Текстовый файл',
    'text/csv': 'CSV файл'
  };
  
  return typeMap[mimeType] || mimeType;
}

/**
 * Сохраняет результаты в таблицу
 */
function saveResults(results) {
  if (results.length === 0) return;
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(CONFIG.SHEET_NAME);
  }
  
  var lastRow = sheet.getLastRow();
  var data = results.map(function(item) {
    return [
      item.name,
      item.mimeType,
      item.url,
      item.owner,
      getAccessTypeDescription(
        DriveApp.Access[item.sharingAccess],
        DriveApp.Permission[item.sharingPermission]
      ),
      item.sharingAccess + ' / ' + item.sharingPermission,
      item.date
    ];
  });
  
  if (data.length > 0) {
    sheet.getRange(lastRow + 1, 1, data.length, data[0].length).setValues(data);
    
    // Автоматически подгоняем ширину колонок (только для первой порции)
    if (lastRow === 1) {
      for (var i = 1; i <= 5; i++) {
        sheet.autoResizeColumn(i);
      }
    }
  }
}

/**
 * Записывает сообщение в лог
 */
function logMessage(message) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    
    if (!logSheet) {
      logSheet = ss.insertSheet(CONFIG.LOG_SHEET_NAME);
      logSheet.getRange(1, 1, 1, 2).setValues([['Время', 'Сообщение']]);
      logSheet.getRange(1, 1, 1, 2).setFontWeight('bold');
    }
    
    var lastRow = logSheet.getLastRow();
    logSheet.getRange(lastRow + 1, 1, 1, 2).setValues([[new Date(), message]]);
    
    // Также выводим в консоль
    console.log(message);
  } catch (e) {
    console.error('Ошибка логирования: ' + e.message);
  }
}

/**
 * Очищает результаты и логи
 */
function clearResults() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Очистить результаты?',
    'Это удалит все данные из листов результатов и логов.',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response == ui.Button.OK) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    var resultsSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (resultsSheet) {
      resultsSheet.clear();
    }
    
    var logSheet = ss.getSheetByName(CONFIG.LOG_SHEET_NAME);
    if (logSheet) {
      logSheet.clear();
    }
    
    ui.alert('Результаты очищены');
  }
}

/**
 * Сбрасывает прогресс сканирования
 */
function resetProgress() {
  var properties = PropertiesService.getScriptProperties();
  properties.deleteProperty(CONFIG.PROPERTY_KEY);
  
  // Удаляем все триггеры continueScan
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'continueScan') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  SpreadsheetApp.getUi().alert('Прогресс сброшен. Все триггеры удалены.');
  logMessage('Прогресс сброшен вручную');
}

/**
 * Показывает текущий статус сканирования
 */
function showStatus() {
  var properties = PropertiesService.getScriptProperties();
  var progress = properties.getProperty(CONFIG.PROPERTY_KEY);
  
  if (progress) {
    var scanState = JSON.parse(progress);
    var message = 'Сканирование в процессе:\n\n' +
                  'Обработано файлов: ' + scanState.processedCount + '\n' +
                  'Найдено проблем: ' + scanState.foundCount + '\n' +
                  'Папок в очереди: ' + scanState.folderQueue.length;
    SpreadsheetApp.getUi().alert('Статус сканирования', message, SpreadsheetApp.getUi().ButtonSet.OK);
  } else {
    SpreadsheetApp.getUi().alert('Сканирование не запущено');
  }
}

/**
 * Функция для ручного запуска полного сканирования
 * (можно использовать из редактора скриптов)
 */
function runFullScan() {
  initializeSheets();
  logMessage('Запуск полного сканирования...');
  
  var allFiles = [];
  var processedCount = 0;
  var foundCount = 0;
  
  // Рекурсивная функция для обхода папок
  function scanFolder(folder, depth) {
    if (depth > 20) {
      logMessage('Достигнута максимальная глубина вложенности для папки: ' + folder.getName());
      return;
    }
    
    try {
      var fileResults = scanFilesInFolder(folder);
      allFiles = allFiles.concat(fileResults.files);
      processedCount += fileResults.processed;
      foundCount += fileResults.found;
      
      // Сохраняем порциями
      if (allFiles.length >= CONFIG.BATCH_SIZE) {
        saveResults(allFiles);
        allFiles = [];
      }
      
      var subfolders = folder.getFolders();
      while (subfolders.hasNext()) {
        scanFolder(subfolders.next(), depth + 1);
      }
    } catch (e) {
      logMessage('Ошибка при сканировании папки ' + folder.getName() + ': ' + e.message);
    }
  }
  
  scanFolder(DriveApp.getRootFolder(), 0);
  
  // Сохраняем последние результаты
  if (allFiles.length > 0) {
    saveResults(allFiles);
  }
  
  logMessage('✅ Полное сканирование завершено! Обработано: ' + processedCount + ', Найдено: ' + foundCount);
}

/**
 * Тестовая функция для проверки работы скрипта
 */
function testScript() {
  logMessage('=== Тестирование скрипта ===');
  
  try {
    // Тест 1: Проверка доступа к Drive
    var root = DriveApp.getRootFolder();
    logMessage('✓ Доступ к Drive получен');
    
    // Тест 2: Проверка листа
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    logMessage('✓ Доступ к таблице получен: ' + ss.getName());
    
    // Тест 3: Проверка Properties Service
    var properties = PropertiesService.getScriptProperties();
    properties.setProperty('test', 'value');
    properties.deleteProperty('test');
    logMessage('✓ Properties Service работает');
    
    // Тест 4: Сканирование первых 5 файлов
    var files = root.getFiles();
    var count = 0;
    while (files.hasNext() && count < 5) {
      var file = files.next();
      logMessage('Файл: ' + file.getName() + ' | Доступ: ' + file.getSharingAccess());
      count++;
    }
    
    logMessage('=== Тест завершен успешно ===');
    SpreadsheetApp.getUi().alert('Тест пройден успешно! Проверьте лист "Логи".');
    
  } catch (e) {
    logMessage('❌ Ошибка теста: ' + e.message);
    SpreadsheetApp.getUi().alert('Ошибка: ' + e.message);
  }
}

/**
 * Удаляет все триггеры (полезно для очистки)
 */
function deleteAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
  logMessage('Все триггеры удалены');
  SpreadsheetApp.getUi().alert('Все триггеры удалены');
}
