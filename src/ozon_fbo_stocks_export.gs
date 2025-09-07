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
    .addItem('🚀 Тест v4 API с пагинацией', 'testV4Pagination')
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
      .addItem('🔬 Анализ v3 API', 'analyzeV3Response')
      .addItem('🔬 Анализ v4 API', 'analyzeV4Response')
      .addItem('📋 Показать настройки', 'showCurrentSettings'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📄 Управление листами')
      .addItem('📋 Список листов магазинов', 'showStoreSheets')
      .addItem('🗑️ Удалить листы магазинов', 'deleteStoreSheets')
      .addItem('🔄 Переименовать листы', 'renameStoreSheets'))
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
    
    // Используем новый метод с пагинацией v4 API
    let allStocks = fetchAllFboStocksV4();
    
    if (allStocks.length === 0) {
      console.log('v4 API не вернул данные, пробуем v3...');
      allStocks = getFBOStocksV3();
      
      if (allStocks.length === 0) {
        console.log('v3 API не вернул данные, пробуем аналитику...');
        allStocks = getFBOStocksAnalytics();
      }
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
    let fboStocks = getFBOStocksV3(warehouseIds);
    
    if (fboStocks.length === 0) {
      console.log('v3 API не вернул данные, пробуем аналитику...');
      fboStocks = getFBOStocksAnalytics(warehouseIds);
    }
    
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
 * Получает все FBO остатки с пагинацией через v4 API
 */
function fetchAllFboStocksV4() {
  const config = getOzonConfig();
  const headers = {
    'Client-Id': config.CLIENT_ID,
    'Api-Key': config.API_KEY
  };
  
  let lastId = '';
  const result = [];
  let pageCount = 0;
  const PAGE_LIMIT = 1000;

  console.log('Начинаем пагинацию по v4 API...');

  do {
    pageCount++;
    console.log(`Обрабатываем страницу ${pageCount}...`);
    
    const payload = {
      filter: {
        visibility: 'ALL' // берём все видимости
      },
      limit: PAGE_LIMIT,
      last_id: lastId
    };

    const resp = callOzonAPI('/v4/product/info/stocks', payload, headers);

    // Обрабатываем разные варианты структуры ответа
    let items = [];
    if (resp && resp.result && Array.isArray(resp.result.items)) {
      items = resp.result.items;
      lastId = resp.result.last_id || '';
    } else if (resp && Array.isArray(resp.items)) {
      items = resp.items;
      lastId = resp.last_id || '';
    } else if (resp && Array.isArray(resp.result)) {
      items = resp.result;
      lastId = resp.last_id || '';
    } else {
      items = [];
      lastId = '';
    }

    console.log(`Получено ${items.length} товаров на странице ${pageCount}`);

    // Трансформируем: оставляем только FBO остатки
    for (const it of items) {
      const productId = it.product_id || it.id || '';
      const offerId = it.offer_id || '';
      const sku = it.sku || '';
      const stocks = Array.isArray(it.stocks) ? it.stocks : [];
      const fbo = stocks.find(s => (s.type || '').toLowerCase() === 'fbo');

      // В v4 иногда добавляют детали склада: s.warehouse_ids (массив).
      const warehouseIds = fbo && Array.isArray(fbo.warehouse_ids) ? fbo.warehouse_ids.join(',') : '';

      if (fbo) { // Добавляем только товары с FBO остатками
        result.push({
          product_id: productId,
          offer_id: offerId,
          sku: sku,
          name: it.name || '',
          fbo_present: fbo ? Number(fbo.present || 0) : 0,
          fbo_reserved: fbo ? Number(fbo.reserved || 0) : 0,
          warehouse_ids: warehouseIds,
          store_name: config.STORE_NAME || 'Неизвестный магазин'
        });
      }
    }

  } while (lastId);

  console.log(`Пагинация завершена. Всего страниц: ${pageCount}, FBO товаров: ${result.length}`);
  return result;
}

/**
 * Низкоуровневый запрос к Ozon Seller API
 */
function callOzonAPI(path, body, headers) {
  const config = getOzonConfig();
  const url = config.BASE_URL + path;
  
  const resp = UrlFetchApp.fetch(url, {
    method: 'post',
    muteHttpExceptions: true,
    contentType: 'application/json; charset=utf-8',
    headers: headers,
    payload: JSON.stringify(body)
  });

  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`Ozon API ${path} вернул код ${code}: ${resp.getContentText()}`);
  }
  const text = resp.getContentText();
  return text ? JSON.parse(text) : {};
}

/**
 * Получает остатки товаров через v3 API (резервный метод)
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
 * Очищает название листа от недопустимых символов
 */
function sanitizeSheetName(name) {
  // Google Sheets ограничения: максимум 100 символов, нельзя использовать: \ / ? * [ ]
  let cleanName = name
    .replace(/[\\\/\?\*\[\]]/g, '') // Удаляем недопустимые символы
    .replace(/\s+/g, ' ') // Заменяем множественные пробелы на один
    .trim(); // Убираем пробелы в начале и конце
  
  // Ограничиваем длину до 100 символов
  if (cleanName.length > 100) {
    cleanName = cleanName.substring(0, 100);
  }
  
  // Если название пустое, используем дефолтное
  if (!cleanName) {
    cleanName = 'FBO Stocks';
  }
  
  return cleanName;
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
  
  // Определяем название листа на основе магазина
  const storeName = config.STORE_NAME || 'Неизвестный магазин';
  const sheetName = sanitizeSheetName(storeName);
  
  console.log(`Создаем/используем лист: ${sheetName}`);
  
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  // Создаем лист если не существует
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  // Очищаем только диапазон с данными (A:J)
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    const range = sheet.getRange(1, 1, lastRow, 10); // 10 колонок A-J
    range.clear();
  }
  
  // Заголовки
  const headers = [
    'Магазин',
    'Product ID',
    'SKU',
    'Название товара',
    'Артикул',
    'FBO Остаток',
    'FBO Зарезервировано',
    'FBO Доступно',
    'ID складов',
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
    // Проверяем структуру данных - новый v4 API или старые методы
    if (stock.fbo_present !== undefined) {
      // Новый v4 API с пагинацией - уже обработанные данные
      rows.push([
        stock.store_name || config.STORE_NAME || 'Неизвестный магазин',
        stock.product_id || '',
        stock.sku || '',
        stock.name || '',
        stock.offer_id || '',
        stock.fbo_present || 0,
        stock.fbo_reserved || 0,
        (stock.fbo_present || 0) - (stock.fbo_reserved || 0), // available = present - reserved
        stock.warehouse_ids || '',
        new Date().toLocaleString('ru-RU')
      ]);
    } else if (stock.stocks && Array.isArray(stock.stocks)) {
      // v3 API - у товара есть массив stocks
      stock.stocks.forEach(stockItem => {
        // Показываем только FBO остатки
        if (stockItem.type === 'fbo') {
          rows.push([
            stock.store_name || config.STORE_NAME || 'Неизвестный магазин',
            stock.product_id || '',
            stock.sku || '',
            stock.name || '',
            stock.offer_id || '',
            stockItem.present || 0,
            stockItem.reserved || 0,
            (stockItem.present || 0) - (stockItem.reserved || 0), // available = present - reserved
            stockItem.warehouse_ids && stockItem.warehouse_ids.length > 0 ? stockItem.warehouse_ids.join(',') : '',
            new Date().toLocaleString('ru-RU')
          ]);
        }
      });
    } else if (stock.available_stock_count !== undefined) {
      // API аналитики - полная структура
      rows.push([
        stock.store_name || config.STORE_NAME || 'Неизвестный магазин',
        stock.product_id || '',
        stock.sku || '',
        stock.name || '',
        stock.offer_id || '',
        stock.available_stock_count || 0,
        0, // reserved не доступен в аналитике
        stock.available_stock_count || 0,
        stock.warehouse_id || '',
        new Date().toLocaleString('ru-RU')
      ]);
    } else {
      // Старая структура API
      rows.push([
        stock.store_name || config.STORE_NAME || 'Неизвестный магазин',
        stock.product_id || '',
        stock.sku || '',
        stock.name || '',
        stock.offer_id || '',
        stock.present || 0,
        stock.reserved || 0,
        stock.available || 0,
        stock.warehouse_id || '',
        new Date().toLocaleString('ru-RU')
      ]);
    }
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
    
    const originalActiveStore = getActiveStore();
    let totalProcessed = 0;
    
    stores.forEach((store, index) => {
      try {
        console.log(`Обрабатываем магазин ${index + 1}/${stores.length}: ${store.name}`);
        
        // Устанавливаем активный магазин
        setActiveStore(store.id);
        
        // Получаем остатки для текущего магазина
        let storeStocks = fetchAllFboStocksV4();
        
        if (storeStocks.length === 0) {
          console.log(`  v4 API не вернул данные, пробуем v3...`);
          storeStocks = getFBOStocksV3();
          
          if (storeStocks.length === 0) {
            console.log(`  v3 API не вернул данные, пробуем аналитику...`);
            storeStocks = getFBOStocksAnalytics();
          }
        }
        
        // Добавляем название магазина к каждому товару
        storeStocks.forEach(stock => {
          stock.store_name = store.name;
        });
        
        console.log(`  Получено ${storeStocks.length} товаров для магазина "${store.name}"`);
        
        // Записываем в отдельный лист для этого магазина
        if (storeStocks.length > 0) {
          writeToGoogleSheets(storeStocks);
          totalProcessed += storeStocks.length;
        }
        
        console.log(`  Магазин "${store.name}" обработан успешно`);
        
      } catch (error) {
        console.error(`Ошибка при обработке магазина "${store.name}":`, error);
        // Продолжаем с другими магазинами
      }
    });
    
    // Восстанавливаем активный магазин
    if (originalActiveStore) {
      setActiveStore(originalActiveStore.id);
    }
    
    console.log(`Выгрузка со всех магазинов завершена! Всего обработано товаров: ${totalProcessed}`);
    
    SpreadsheetApp.getUi().alert('Выгрузка завершена', `Обработано ${stores.length} магазинов, всего товаров: ${totalProcessed}`, SpreadsheetApp.getUi().ButtonSet.OK);
    
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
 * Показывает список листов магазинов
 */
function showStoreSheets() {
  const config = getOzonConfig();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  
  const storeSheets = sheets.filter(sheet => {
    const sheetName = sheet.getName();
    const stores = getStoresList();
    return stores.some(store => sanitizeSheetName(store.name) === sheetName);
  });
  
  if (storeSheets.length === 0) {
    SpreadsheetApp.getUi().alert('Информация', 'Нет листов магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  let message = 'Листы магазинов:\n\n';
  storeSheets.forEach((sheet, index) => {
    const rowCount = sheet.getLastRow() - 1; // -1 для заголовка
    message += `${index + 1}. ${sheet.getName()} (${rowCount} товаров)\n`;
  });
  
  SpreadsheetApp.getUi().alert('Листы магазинов', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Удаляет листы магазинов
 */
function deleteStoreSheets() {
  const ui = SpreadsheetApp.getUi();
  const confirm = ui.alert('Подтверждение', 'Удалить все листы магазинов?', ui.ButtonSet.YES_NO);
  
  if (confirm === ui.Button.YES) {
    const config = getOzonConfig();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = spreadsheet.getSheets();
    const stores = getStoresList();
    
    let deletedCount = 0;
    
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const isStoreSheet = stores.some(store => sanitizeSheetName(store.name) === sheetName);
      
      if (isStoreSheet) {
        spreadsheet.deleteSheet(sheet);
        deletedCount++;
      }
    });
    
    ui.alert('Успех', `Удалено листов: ${deletedCount}`, ui.ButtonSet.OK);
  }
}

/**
 * Переименовывает листы магазинов
 */
function renameStoreSheets() {
  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = spreadsheet.getSheets();
  const stores = getStoresList();
  
  let renamedCount = 0;
  
  stores.forEach(store => {
    const expectedSheetName = sanitizeSheetName(store.name);
    const existingSheet = spreadsheet.getSheetByName(expectedSheetName);
    
    if (!existingSheet) {
      // Ищем лист с неправильным названием
      const oldSheet = sheets.find(sheet => {
        const sheetName = sheet.getName();
        return sheetName.includes(store.name) || store.name.includes(sheetName);
      });
      
      if (oldSheet && oldSheet.getName() !== expectedSheetName) {
        try {
          oldSheet.setName(expectedSheetName);
          renamedCount++;
        } catch (error) {
          console.error(`Ошибка переименования листа ${oldSheet.getName()}:`, error);
        }
      }
    }
  });
  
  ui.alert('Переименование завершено', `Переименовано листов: ${renamedCount}`, ui.ButtonSet.OK);
}

/**
 * Тестирует v4 API с пагинацией
 */
function testV4Pagination() {
  try {
    const config = getOzonConfig();
    if (!config.CLIENT_ID || !config.API_KEY) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроены API ключи!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log('Начинаем тест v4 API с пагинацией...');
    
    const result = fetchAllFboStocksV4();
    
    console.log(`Тест завершен. Получено FBO товаров: ${result.length}`);
    
    if (result.length > 0) {
      console.log('Примеры данных:');
      result.slice(0, 3).forEach((item, index) => {
        console.log(`${index + 1}. ${item.offer_id} - FBO: ${item.fbo_present}, Reserved: ${item.fbo_reserved}`);
      });
    }
    
    SpreadsheetApp.getUi().alert('Тест завершен', `Получено FBO товаров: ${result.length}`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Ошибка теста v4 API:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка теста: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Детально анализирует ответ от v3 API
 */
function analyzeV3Response() {
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
  console.log(`Анализируем ответ v3 API для склада: ${testWarehouseId}`);
  
  try {
    const url = `${config.BASE_URL}/v3/product/info/stocks`;
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
      console.log('📋 Полная структура ответа v3:');
      console.log(JSON.stringify(data, null, 2));
    }
    
  } catch (error) {
    console.error('Ошибка анализа v3 API:', error);
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
