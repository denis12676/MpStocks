/**
 * Скрипт для выгрузки остатков со складов FBO Ozon Seller в Google Таблицы
 * Требует настройки API ключей Ozon и Google Sheets
 */

// Конфигурация API Ozon
const OZON_CONFIG = {
  BASE_URL: 'https://api-seller.ozon.ru'
};

/**
 * Получает настройки из PropertiesService
 */
function getOzonConfig() {
  const properties = PropertiesService.getScriptProperties();
  
  return {
    CLIENT_ID: properties.getProperty('OZON_CLIENT_ID'),
    API_KEY: properties.getProperty('OZON_API_KEY'),
    SPREADSHEET_ID: properties.getProperty('GOOGLE_SPREADSHEET_ID'),
    BASE_URL: OZON_CONFIG.BASE_URL
  };
}

/**
 * Сохраняет настройки в PropertiesService
 */
function saveOzonConfig(clientId, apiKey, spreadsheetId) {
  const properties = PropertiesService.getScriptProperties();
  
  properties.setProperties({
    'OZON_CLIENT_ID': clientId,
    'OZON_API_KEY': apiKey,
    'GOOGLE_SPREADSHEET_ID': spreadsheetId
  });
  
  console.log('Настройки сохранены успешно!');
}

/**
 * Показывает текущие настройки (без API ключей)
 */
function showCurrentSettings() {
  const config = getOzonConfig();
  
  console.log('Текущие настройки:');
  console.log('Client ID:', config.CLIENT_ID ? '***' + config.CLIENT_ID.slice(-4) : 'Не установлен');
  console.log('API Key:', config.API_KEY ? '***' + config.API_KEY.slice(-4) : 'Не установлен');
  console.log('Spreadsheet ID:', config.SPREADSHEET_ID || 'Не установлен');
}

/**
 * Создает меню при открытии Google Таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🛒 Ozon FBO Export')
    .addItem('📊 Выгрузить остатки', 'exportFBOStocks')
    .addSeparator()
    .addSubMenu(ui.createMenu('🏪 Управление магазинами')
      .addItem('➕ Добавить магазин', 'addNewStore')
      .addItem('📋 Список магазинов', 'showStoresList')
      .addItem('✏️ Редактировать магазин', 'editStore')
      .addItem('🗑️ Удалить магазин', 'deleteStore')
      .addItem('🔄 Переключить активный магазин', 'switchActiveStore'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Настройки')
      .addItem('📊 ID Google Таблицы', 'setSpreadsheetId')
      .addItem('🔍 Проверить подключение', 'testOzonConnection')
      .addItem('📋 Показать настройки', 'showCurrentSettings'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⏰ Автоматизация')
      .addItem('🕘 Создать ежедневный триггер', 'createDailyTrigger')
      .addItem('❌ Удалить все триггеры', 'deleteAllTriggers'))
    .addToUi();
}

/**
 * Получает список всех магазинов
 */
function getStoresList() {
  const properties = PropertiesService.getScriptProperties();
  const storesJson = properties.getProperty('OZON_STORES');
  
  if (!storesJson) {
    return [];
  }
  
  try {
    return JSON.parse(storesJson);
  } catch (error) {
    console.error('Ошибка парсинга списка магазинов:', error);
    return [];
  }
}

/**
 * Сохраняет список магазинов
 */
function saveStoresList(stores) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('OZON_STORES', JSON.stringify(stores));
}

/**
 * Получает активный магазин
 */
function getActiveStore() {
  const properties = PropertiesService.getScriptProperties();
  const activeStoreId = properties.getProperty('ACTIVE_STORE_ID');
  
  if (!activeStoreId) {
    return null;
  }
  
  const stores = getStoresList();
  return stores.find(store => store.id === activeStoreId) || null;
}

/**
 * Устанавливает активный магазин
 */
function setActiveStore(storeId) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('ACTIVE_STORE_ID', storeId);
}

/**
 * Добавляет новый магазин
 */
function addNewStore() {
  const ui = SpreadsheetApp.getUi();
  
  // Запрашиваем данные магазина
  const storeName = ui.prompt('Добавить магазин', 'Введите название магазина:', ui.ButtonSet.OK_CANCEL);
  if (storeName.getSelectedButton() !== ui.Button.OK) return;
  
  const clientId = ui.prompt('Добавить магазин', 'Введите Client ID:', ui.ButtonSet.OK_CANCEL);
  if (clientId.getSelectedButton() !== ui.Button.OK) return;
  
  const apiKey = ui.prompt('Добавить магазин', 'Введите API Key:', ui.ButtonSet.OK_CANCEL);
  if (apiKey.getSelectedButton() !== ui.Button.OK) return;
  
  // Создаем новый магазин
  const newStore = {
    id: Utilities.getUuid(),
    name: storeName.getResponseText(),
    clientId: clientId.getResponseText(),
    apiKey: apiKey.getResponseText(),
    createdAt: new Date().toISOString()
  };
  
  // Добавляем в список
  const stores = getStoresList();
  stores.push(newStore);
  saveStoresList(stores);
  
  // Если это первый магазин, делаем его активным
  if (stores.length === 1) {
    setActiveStore(newStore.id);
  }
  
  ui.alert('Успех', `Магазин "${newStore.name}" добавлен!`, ui.ButtonSet.OK);
}

/**
 * Показывает список магазинов
 */
function showStoresList() {
  const stores = getStoresList();
  const activeStore = getActiveStore();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Список магазинов пуст', 'Добавьте магазины через меню "Управление магазинами"', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  let message = 'Список магазинов:\n\n';
  stores.forEach((store, index) => {
    const isActive = activeStore && store.id === activeStore.id ? ' (АКТИВНЫЙ)' : '';
    message += `${index + 1}. ${store.name}${isActive}\n`;
    message += `   Client ID: ***${store.clientId.slice(-4)}\n`;
    message += `   API Key: ***${store.apiKey.slice(-4)}\n\n`;
  });
  
  SpreadsheetApp.getUi().alert('Список магазинов', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Переключает активный магазин
 */
function switchActiveStore() {
  const stores = getStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (stores.length === 1) {
    SpreadsheetApp.getUi().alert('Информация', 'У вас только один магазин', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите магазин:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Переключить магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const selectedStore = stores[selectedIndex];
    setActiveStore(selectedStore.id);
    ui.alert('Успех', `Активный магазин изменен на: ${selectedStore.name}`, ui.ButtonSet.OK);
  } else {
    ui.alert('Ошибка', 'Неверный номер магазина', ui.ButtonSet.OK);
  }
}

/**
 * Редактирует магазин
 */
function editStore() {
  const stores = getStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите магазин для редактирования:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Редактировать магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const store = stores[selectedIndex];
    
    // Запрашиваем новые данные
    const newName = ui.prompt('Редактировать магазин', `Название (текущее: ${store.name}):`, ui.ButtonSet.OK_CANCEL);
    if (newName.getSelectedButton() !== ui.Button.OK) return;
    
    const newClientId = ui.prompt('Редактировать магазин', `Client ID (текущий: ***${store.clientId.slice(-4)}):`, ui.ButtonSet.OK_CANCEL);
    if (newClientId.getSelectedButton() !== ui.Button.OK) return;
    
    const newApiKey = ui.prompt('Редактировать магазин', `API Key (текущий: ***${store.apiKey.slice(-4)}):`, ui.ButtonSet.OK_CANCEL);
    if (newApiKey.getSelectedButton() !== ui.Button.OK) return;
    
    // Обновляем данные
    store.name = newName.getResponseText();
    store.clientId = newClientId.getResponseText();
    store.apiKey = newApiKey.getResponseText();
    
    saveStoresList(stores);
    ui.alert('Успех', `Магазин "${store.name}" обновлен!`, ui.ButtonSet.OK);
  } else {
    ui.alert('Ошибка', 'Неверный номер магазина', ui.ButtonSet.OK);
  }
}

/**
 * Удаляет магазин
 */
function deleteStore() {
  const stores = getStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите магазин для удаления:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Удалить магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const storeToDelete = stores[selectedIndex];
    const confirm = ui.alert('Подтверждение', `Удалить магазин "${storeToDelete.name}"?`, ui.ButtonSet.YES_NO);
    
    if (confirm === ui.Button.YES) {
      stores.splice(selectedIndex, 1);
      saveStoresList(stores);
      
      // Если удалили активный магазин, выбираем новый
      const activeStore = getActiveStore();
      if (!activeStore && stores.length > 0) {
        setActiveStore(stores[0].id);
      }
      
      ui.alert('Успех', 'Магазин удален!', ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Ошибка', 'Неверный номер магазина', ui.ButtonSet.OK);
  }
}

/**
 * Устанавливает ID Google Таблицы
 */
function setSpreadsheetId() {
  const ui = SpreadsheetApp.getUi();
  const currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  const response = ui.prompt('ID Google Таблицы', `Текущий ID: ${currentId}\n\nВведите новый ID (или оставьте пустым для текущего):`, ui.ButtonSet.OK_CANCEL);
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const newId = response.getResponseText().trim();
    const spreadsheetId = newId || currentId;
    
    const properties = PropertiesService.getScriptProperties();
    properties.setProperty('GOOGLE_SPREADSHEET_ID', spreadsheetId);
    
    ui.alert('Успех', `ID таблицы установлен: ${spreadsheetId}`, ui.ButtonSet.OK);
  }
}

/**
 * Основная функция для запуска выгрузки
 */
function exportFBOStocks() {
  try {
    console.log('Начинаем выгрузку остатков FBO...');
    
    // Проверяем настройки
    const config = getOzonConfig();
    if (!config.CLIENT_ID || !config.API_KEY) {
      throw new Error('Не настроены API ключи! Используйте saveOzonConfig() для настройки.');
    }
    
    // Получаем данные о складах
    const warehouses = getWarehouses();
    console.log(`Найдено складов: ${warehouses.length}`);
    
    // Получаем остатки по всем складам
    const allStocks = [];
    warehouses.forEach(warehouse => {
      const stocks = getFBOStocks(warehouse.warehouse_id);
      stocks.forEach(stock => {
        stock.warehouse_name = warehouse.name;
        stock.warehouse_id = warehouse.warehouse_id;
      });
      allStocks.push(...stocks);
    });
    
    console.log(`Получено записей об остатках: ${allStocks.length}`);
    
    // Записываем в Google Таблицы
    writeToGoogleSheets(allStocks);
    
    console.log('Выгрузка завершена успешно!');
    
  } catch (error) {
    console.error('Ошибка при выгрузке:', error);
    throw error;
  }
}

/**
 * Получает список складов FBO
 */
function getWarehouses() {
  const config = getOzonConfig();
  const url = `${config.BASE_URL}/v1/warehouse/list`;
  
  const options = {
    method: 'POST',
    headers: {
      'Client-Id': config.CLIENT_ID,
      'Api-Key': config.API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({
      filter: {
        type: ['FBO'] // Только FBO склады
      }
    })
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  
  if (!data.result) {
    throw new Error('Ошибка получения списка складов: ' + JSON.stringify(data));
  }
  
  return data.result;
}

/**
 * Получает остатки товаров на конкретном складе FBO
 */
function getFBOStocks(warehouseId) {
  const config = getOzonConfig();
  // Пробуем разные версии API
  const apiVersions = ['/v2/product/info/stocks', '/v1/product/info/stocks', '/v3/product/info/stocks'];
  
  for (let i = 0; i < apiVersions.length; i++) {
    try {
      const url = `${config.BASE_URL}${apiVersions[i]}`;
      
      const options = {
        method: 'POST',
        headers: {
          'Client-Id': config.CLIENT_ID,
          'Api-Key': config.API_KEY,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          filter: {
            warehouse_id: [warehouseId]
          },
          limit: 1000
        })
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const data = JSON.parse(response.getContentText());
      
      if (data.result && data.result.items) {
        console.log(`Успешно получены остатки через ${apiVersions[i]}`);
        return data.result.items;
      }
      
    } catch (error) {
      console.log(`Ошибка с ${apiVersions[i]}: ${error.message}`);
      if (i === apiVersions.length - 1) {
        throw new Error(`Все версии API недоступны. Последняя ошибка: ${error.message}`);
      }
    }
  }
  
  return [];
}

/**
 * Записывает данные в Google Таблицы
 */
function writeToGoogleSheets(stocks) {
  const config = getOzonConfig();
  const spreadsheet = SpreadsheetApp.openById(config.SPREADSHEET_ID);
  let sheet = spreadsheet.getSheetByName('FBO Stocks');
  
  // Создаем лист если не существует
  if (!sheet) {
    sheet = spreadsheet.insertSheet('FBO Stocks');
  }
  
  // Очищаем лист
  sheet.clear();
  
  // Заголовки
  const headers = [
    'SKU',
    'Название товара',
    'Артикул',
    'ID склада',
    'Название склада',
    'Остаток',
    'Зарезервировано',
    'Доступно',
    'Дата обновления'
  ];
  
  // Записываем заголовки
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Форматируем заголовки
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#E8F0FE');
  
  if (stocks.length === 0) {
    console.log('Нет данных для записи');
    return;
  }
  
  // Подготавливаем данные
  const rows = stocks.map(stock => [
    stock.offer_id || '',
    stock.name || '',
    stock.article || '',
    stock.warehouse_id || '',
    stock.warehouse_name || '',
    stock.present || 0,
    stock.reserved || 0,
    stock.available || 0,
    new Date().toLocaleString('ru-RU')
  ]);
  
  // Записываем данные
  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }
  
  // Автоподбор ширины колонок
  sheet.autoResizeColumns(1, headers.length);
  
  // Добавляем фильтры
  sheet.getRange(1, 1, rows.length + 1, headers.length).createFilter();
  
  console.log(`Записано ${rows.length} строк в Google Таблицы`);
}

/**
 * Функция для тестирования подключения к API Ozon
 */
function testOzonConnection() {
  try {
    const config = getOzonConfig();
    if (!config.CLIENT_ID || !config.API_KEY) {
      console.error('Не настроены API ключи! Используйте saveOzonConfig() для настройки.');
      return false;
    }
    
    const warehouses = getWarehouses();
    console.log('Подключение к Ozon API успешно!');
    console.log('Найденные склады:', warehouses);
    return true;
  } catch (error) {
    console.error('Ошибка подключения к Ozon API:', error);
    return false;
  }
}

/**
 * Создает триггер для автоматического запуска (ежедневно в 9:00)
 */
function createDailyTrigger() {
  // Удаляем существующие триггеры
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'exportFBOStocks') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Создаем новый триггер
  ScriptApp.newTrigger('exportFBOStocks')
    .timeBased()
    .everyDays(1)
    .atHour(9)
    .create();
    
  console.log('Триггер для ежедневного запуска создан (9:00)');
}

/**
 * Удаляет все триггеры
 */
function deleteAllTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    ScriptApp.deleteTrigger(trigger);
  });
  console.log('Все триггеры удалены');
}
