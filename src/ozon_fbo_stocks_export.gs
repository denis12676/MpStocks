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
  
  // Сначала пробуем получить из активного магазина
  const activeStore = getActiveStore();
  
  if (activeStore) {
    return {
      CLIENT_ID: activeStore.clientId,
      API_KEY: activeStore.apiKey,
      SPREADSHEET_ID: properties.getProperty('GOOGLE_SPREADSHEET_ID'),
      BASE_URL: OZON_CONFIG.BASE_URL,
      STORE_NAME: activeStore.name
    };
  }
  
  // Fallback на старые настройки
  return {
    CLIENT_ID: properties.getProperty('OZON_CLIENT_ID'),
    API_KEY: properties.getProperty('OZON_API_KEY'),
    SPREADSHEET_ID: properties.getProperty('GOOGLE_SPREADSHEET_ID'),
    BASE_URL: OZON_CONFIG.BASE_URL,
    STORE_NAME: 'Legacy Store'
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
    .addItem('📊 Выгрузить все остатки (активный магазин)', 'exportFBOStocks')
    .addItem('📊 Выгрузить только FBO остатки', 'exportOnlyFBOStocks')
    .addItem('📊 Выгрузить остатки (все магазины)', 'exportAllStoresStocks')
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
      .addItem('📊 Установить текущую таблицу', 'setCurrentSpreadsheetId')
      .addItem('🔍 Проверить подключение', 'testOzonConnection')
      .addItem('🧪 Тест API endpoints', 'testStocksEndpoints')
      .addItem('🔬 Анализ v4 API', 'analyzeV4Response')
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
    
    // Проверяем валидность ID
    try {
      SpreadsheetApp.openById(spreadsheetId);
      const properties = PropertiesService.getScriptProperties();
      properties.setProperty('GOOGLE_SPREADSHEET_ID', spreadsheetId);
      ui.alert('Успех', `ID таблицы установлен: ${spreadsheetId}`, ui.ButtonSet.OK);
    } catch (error) {
      ui.alert('Ошибка', `Неверный ID таблицы: ${error.message}`, ui.ButtonSet.OK);
    }
  }
}

/**
 * Автоматически устанавливает ID текущей таблицы
 */
function setCurrentSpreadsheetId() {
  const ui = SpreadsheetApp.getUi();
  const currentId = SpreadsheetApp.getActiveSpreadsheet().getId();
  
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('GOOGLE_SPREADSHEET_ID', currentId);
  
  ui.alert('Успех', `ID текущей таблицы установлен: ${currentId}`, ui.ButtonSet.OK);
}

/**
 * Основная функция для запуска выгрузки
 */
function exportFBOStocks() {
  try {
    // Проверяем настройки
    const config = getOzonConfig();
    if (!config.CLIENT_ID || !config.API_KEY) {
      throw new Error('Не настроены API ключи! Добавьте магазин через меню "Управление магазинами".');
    }
    
    console.log(`Начинаем выгрузку остатков FBO для магазина: ${config.STORE_NAME}...`);
    
    // Пробуем сначала v3 API, затем аналитику как резерв
    let allStocks = getFBOStocksV3();
    
    if (allStocks.length === 0) {
      console.log('v3 API не вернул данные, пробуем аналитику...');
      allStocks = getFBOStocksAnalytics();
    }
    
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
 * Выгружает только FBO остатки (без FBS)
 */
function exportOnlyFBOStocks() {
  try {
    // Проверяем настройки
    const config = getOzonConfig();
    if (!config.CLIENT_ID || !config.API_KEY) {
      throw new Error('Не настроены API ключи! Добавьте магазин через меню "Управление магазинами".');
    }
    
    console.log(`Начинаем выгрузку только FBO остатков для магазина: ${config.STORE_NAME}...`);
    
    // Получаем только FBO склады
    const fboWarehouses = getWarehouses();
    console.log(`Найдено FBO складов: ${fboWarehouses.length}`);
    
    if (fboWarehouses.length === 0) {
      console.log('Нет FBO складов для выгрузки');
      return;
    }
    
    // Получаем остатки только с FBO складов
    const warehouseIds = fboWarehouses.map(w => w.warehouse_id);
    const fboStocks = getFBOStocksAnalytics(warehouseIds);
    
    console.log(`Получено записей об FBO остатках: ${fboStocks.length}`);
    
    // Записываем в Google Таблицы
    writeToGoogleSheets(fboStocks);
    
    console.log('Выгрузка FBO остатков завершена успешно!');
    
  } catch (error) {
    console.error('Ошибка при выгрузке FBO остатков:', error);
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
 * Получает все склады (FBO и FBS)
 */
function getAllWarehouses() {
  const config = getOzonConfig();
  const url = `${config.BASE_URL}/v1/warehouse/list`;
  
  const options = {
    method: 'POST',
    headers: {
      'Client-Id': config.CLIENT_ID,
      'Api-Key': config.API_KEY,
      'Content-Type': 'application/json'
    },
    payload: JSON.stringify({}) // Без фильтра - получаем все склады
  };
  
  const response = UrlFetchApp.fetch(url, options);
  const data = JSON.parse(response.getContentText());
  
  if (!data.result) {
    throw new Error('Ошибка получения списка складов: ' + JSON.stringify(data));
  }
  
  return data.result;
}

/**
 * Получает остатки товаров через v3 API
 */
function getFBOStocksV3(warehouseIds = []) {
  const config = getOzonConfig();
  
  try {
    const url = `${config.BASE_URL}/v3/product/info/stocks`;
    console.log(`Используем endpoint v3: ${url}`);
    
    const payload = {
      filter: {}
    };
    
    // Если указаны конкретные склады, добавляем фильтр
    if (warehouseIds.length > 0) {
      payload.filter.warehouse_id = warehouseIds;
    }
    
    payload.limit = 1000; // Максимум записей за раз
    
    const options = {
      method: 'POST',
      headers: {
        'Client-Id': config.CLIENT_ID,
        'Api-Key': config.API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`Response code: ${responseCode}`);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      console.log(`📋 Структура ответа v3:`, Object.keys(data));
      
      if (data.result && data.result.items && Array.isArray(data.result.items)) {
        console.log(`✅ Успешно получены остатки через v3: ${data.result.items.length} товаров`);
        return data.result.items;
      } else if (data.items && Array.isArray(data.items)) {
        console.log(`✅ Успешно получены остатки через v3 (items): ${data.items.length} товаров`);
        return data.items;
      } else {
        console.log(`⚠️ Неожиданная структура ответа v3:`, data);
        return [];
      }
    } else {
      console.log(`❌ Ошибка ${responseCode} v3: ${responseText}`);
      return [];
    }
    
  } catch (error) {
    console.log(`❌ Исключение v3: ${error.message}`);
    return [];
  }
}

/**
 * Получает остатки товаров через аналитику FBO (резервный метод)
 */
function getFBOStocksAnalytics(warehouseIds = []) {
  const config = getOzonConfig();
  
  try {
    const url = `${config.BASE_URL}/v1/analytics/stocks`;
    console.log(`Используем endpoint аналитики: ${url}`);
    
    const payload = {
      skus: [] // Получаем все товары
    };
    
    // Если указаны конкретные склады, добавляем фильтр
    if (warehouseIds.length > 0) {
      payload.warehouse_ids = warehouseIds;
    }
    
    const options = {
      method: 'POST',
      headers: {
        'Client-Id': config.CLIENT_ID,
        'Api-Key': config.API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`Response code: ${responseCode}`);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      console.log(`📋 Структура ответа аналитики:`, Object.keys(data));
      
      if (data.items && Array.isArray(data.items)) {
        console.log(`✅ Успешно получены остатки через аналитику: ${data.items.length} товаров`);
        return data.items;
      } else {
        console.log(`⚠️ Нет поля items в ответе аналитики:`, Object.keys(data));
        return [];
      }
    } else {
      console.log(`❌ Ошибка ${responseCode} аналитики: ${responseText}`);
      return [];
    }
    
  } catch (error) {
    console.log(`❌ Исключение аналитики: ${error.message}`);
    return [];
  }
}

/**
 * Получает остатки товаров на конкретном складе (старый метод для совместимости)
 */
function getFBOStocks(warehouseId) {
  // Используем новый метод аналитики
  return getFBOStocksAnalytics([warehouseId]);
}

/**
 * Записывает данные в Google Таблицы
 */
function writeToGoogleSheets(stocks) {
  const config = getOzonConfig();
  
  // Получаем ID таблицы
  let spreadsheetId = config.SPREADSHEET_ID;
  
  // Если ID не установлен, используем текущую таблицу
  if (!spreadsheetId) {
    spreadsheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
    console.log(`Используем текущую таблицу: ${spreadsheetId}`);
  }
  
  console.log(`Открываем таблицу с ID: ${spreadsheetId}`);
  
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    console.error(`Ошибка открытия таблицы с ID ${spreadsheetId}:`, error);
    // Пробуем использовать текущую таблицу
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    console.log('Используем текущую активную таблицу');
  }
  let sheet = spreadsheet.getSheetByName('FBO Stocks');
  
  // Создаем лист если не существует
  if (!sheet) {
    sheet = spreadsheet.insertSheet('FBO Stocks');
  }
  
  // Очищаем лист
  sheet.clear();
  
  // Заголовки
  const headers = [
    'Магазин',
    'SKU',
    'Название товара',
    'Артикул',
    'ID склада',
    'Название склада',
    'Доступно к продаже',
    'Валидный остаток',
    'Излишки',
    'В пути',
    'Брак',
    'Возврат от покупателей',
    'Возврат продавцу',
    'Проверка',
    'Заявки на поставку',
    'Статус ликвидности',
    'Среднесуточные продажи',
    'Дней без продаж',
    'Дней хватит остатка',
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
  const rows = [];
  
  stocks.forEach(stock => {
    // API аналитики возвращает структуру с полной информацией об остатках
    rows.push([
      stock.store_name || config.STORE_NAME || 'Неизвестный магазин',
      stock.sku || '',
      stock.name || '',
      stock.offer_id || '',
      stock.warehouse_id || '',
      stock.warehouse_name || '',
      stock.available_stock_count || 0,
      stock.valid_stock_count || 0,
      stock.excess_stock_count || 0,
      stock.transit_stock_count || 0,
      (stock.stock_defect_stock_count || 0) + (stock.transit_defect_stock_count || 0),
      stock.return_from_customer_stock_count || 0,
      stock.return_to_seller_stock_count || 0,
      stock.other_stock_count || 0,
      stock.requested_stock_count || 0,
      stock.turnover_grade || '',
      stock.ads || 0,
      stock.days_without_sales || 0,
      stock.idc || 0,
      new Date().toLocaleString('ru-RU')
    ]);
  });
  
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
 * Выгружает остатки со всех магазинов
 */
function exportAllStoresStocks() {
  try {
    const stores = getStoresList();
    
    if (stores.length === 0) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Нет добавленных магазинов!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Начинаем выгрузку остатков со всех магазинов (${stores.length} магазинов)...`);
    
    const allStocks = [];
    
    stores.forEach((store, index) => {
      try {
        console.log(`Обрабатываем магазин ${index + 1}/${stores.length}: ${store.name}`);
        
        // Временно устанавливаем активный магазин
        const originalActiveStore = getActiveStore();
        setActiveStore(store.id);
        
        // Получаем данные о складах
        const warehouses = getWarehouses();
        console.log(`  Найдено складов: ${warehouses.length}`);
        
        // Получаем остатки по всем складам
        warehouses.forEach(warehouse => {
          const stocks = getFBOStocks(warehouse.warehouse_id);
          stocks.forEach(stock => {
            stock.warehouse_name = warehouse.name;
            stock.warehouse_id = warehouse.warehouse_id;
            stock.store_name = store.name; // Добавляем название магазина
          });
          allStocks.push(...stocks);
        });
        
        // Восстанавливаем активный магазин
        if (originalActiveStore) {
          setActiveStore(originalActiveStore.id);
        }
        
        console.log(`  Магазин "${store.name}" обработан успешно`);
        
      } catch (error) {
        console.error(`Ошибка при обработке магазина "${store.name}":`, error);
        // Продолжаем с другими магазинами
      }
    });
    
    console.log(`Получено записей об остатках со всех магазинов: ${allStocks.length}`);
    
    // Записываем в Google Таблицы
    writeToGoogleSheets(allStocks);
    
    console.log('Выгрузка со всех магазинов завершена успешно!');
    
  } catch (error) {
    console.error('Ошибка при выгрузке со всех магазинов:', error);
    throw error;
  }
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
 * Детально анализирует ответ от v4 API
 */
function analyzeV4Response() {
  const config = getOzonConfig();
  if (!config.CLIENT_ID || !config.API_KEY) {
    console.error('Не настроены API ключи!');
    return;
  }
  
  const warehouses = getWarehouses();
  if (warehouses.length === 0) {
    console.error('Нет доступных складов для тестирования');
    return;
  }
  
  const testWarehouseId = warehouses[0].warehouse_id;
  console.log(`Анализируем ответ v4 API для склада: ${testWarehouseId}`);
  
  try {
    const url = `${config.BASE_URL}/v4/product/info/stocks`;
    console.log(`URL: ${url}`);
    
    const options = {
      method: 'POST',
      headers: {
        'Client-Id': config.CLIENT_ID,
        'Api-Key': config.API_KEY,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify({
        filter: {
          warehouse_id: [testWarehouseId]
        },
        limit: 10
      }),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`Response code: ${responseCode}`);
    console.log(`Response text: ${responseText}`);
    
    if (responseCode === 200) {
      const data = JSON.parse(responseText);
      console.log('📋 Полная структура ответа:');
      console.log(JSON.stringify(data, null, 2));
    }
    
  } catch (error) {
    console.error('Ошибка анализа v4 API:', error);
  }
}

/**
 * Тестирует все доступные API endpoints для остатков
 */
function testStocksEndpoints() {
  const config = getOzonConfig();
  if (!config.CLIENT_ID || !config.API_KEY) {
    console.error('Не настроены API ключи!');
    return;
  }
  
  const warehouses = getWarehouses();
  if (warehouses.length === 0) {
    console.error('Нет доступных складов для тестирования');
    return;
  }
  
  const testWarehouseId = warehouses[0].warehouse_id;
  console.log(`Тестируем API endpoints для склада: ${testWarehouseId}`);
  
  const apiEndpoints = [
    '/v3/product/info/stocks',
    '/v2/product/info/stocks', 
    '/v1/product/info/stocks',
    '/v4/product/info/stocks',
    '/v1/product/stocks',
    '/v2/product/stocks',
    '/v1/warehouse/stocks',
    '/v2/warehouse/stocks'
  ];
  
  apiEndpoints.forEach(endpoint => {
    try {
      const url = `${config.BASE_URL}${endpoint}`;
      console.log(`\n🔍 Тестируем: ${endpoint}`);
      
      const options = {
        method: 'POST',
        headers: {
          'Client-Id': config.CLIENT_ID,
          'Api-Key': config.API_KEY,
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          filter: {
            warehouse_id: [testWarehouseId]
          },
          limit: 10
        }),
        muteHttpExceptions: true
      };
      
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        console.log(`✅ ${endpoint} - OK (200)`);
        try {
          const data = JSON.parse(responseText);
          console.log(`   Структура ответа:`, Object.keys(data));
          if (data.result) {
            console.log(`   Результат:`, typeof data.result, Array.isArray(data.result) ? `массив из ${data.result.length} элементов` : Object.keys(data.result));
          }
        } catch (parseError) {
          console.log(`   Ошибка парсинга JSON: ${parseError.message}`);
        }
      } else {
        console.log(`❌ ${endpoint} - ${responseCode}: ${responseText.substring(0, 200)}...`);
      }
      
    } catch (error) {
      console.log(`❌ ${endpoint} - Исключение: ${error.message}`);
    }
  });
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
