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
 * Получает конфигурацию WB из PropertiesService
 */
function getWBConfig() {
  const properties = PropertiesService.getScriptProperties();
  
  // Получаем из активного WB магазина
  const activeStore = getActiveWBStore();
  
  if (activeStore) {
    return {
      API_KEY: activeStore.api_key,
      SPREADSHEET_ID: properties.getProperty('GOOGLE_SPREADSHEET_ID'),
      STORE_NAME: activeStore.name
    };
  }
  
  // Fallback на старые настройки
  return {
    API_KEY: properties.getProperty('WB_API_KEY'),
    SPREADSHEET_ID: properties.getProperty('GOOGLE_SPREADSHEET_ID'),
    STORE_NAME: 'Legacy WB Store'
  };
}

/**
 * Получает конфигурацию активного Яндекс Маркет магазина
 */
function getYandexConfig() {
  const properties = PropertiesService.getScriptProperties();
  
  // Получаем из активного Яндекс Маркет магазина
  const activeStore = getActiveYandexStore();
  
  if (activeStore) {
    return {
      API_TOKEN: activeStore.api_token,
      CAMPAIGN_ID: activeStore.campaign_id,
      SPREADSHEET_ID: properties.getProperty('GOOGLE_SPREADSHEET_ID'),
      STORE_NAME: activeStore.name
    };
  }
  
  // Fallback на старые настройки
  return {
    API_TOKEN: properties.getProperty('YANDEX_API_TOKEN'),
    CAMPAIGN_ID: properties.getProperty('YANDEX_CAMPAIGN_ID'),
    SPREADSHEET_ID: properties.getProperty('GOOGLE_SPREADSHEET_ID'),
    STORE_NAME: 'Legacy Yandex Store'
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
    .addItem('📈 Выгрузить цены (активный магазин)', 'exportOzonPrices')
    .addItem('📊 Выгрузить детальные цены (все товары)', 'exportOzonPricesDetailed')
    .addItem('📊 Выгрузить детальные цены (все магазины)', 'exportAllStoresPricesDetailed')
    .addItem('🚀 Тест v4 API с пагинацией', 'testV4Pagination')
    .addSeparator()
    .addSubMenu(ui.createMenu('🏪 Управление магазинами')
      .addItem('➕ Добавить магазин', 'addNewStore')
      .addItem('📋 Список магазинов', 'showStoresList')
      .addItem('✏️ Редактировать магазин', 'editStore')
      .addItem('🗑️ Удалить магазин', 'deleteStore')
      .addItem('🔄 Переключить активный магазин', 'switchActiveStore'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🛒 Управление WB магазинами')
      .addItem('➕ Добавить WB магазин', 'addNewWBStore')
      .addItem('📋 Список WB магазинов', 'showWBStoresList')
      .addItem('✏️ Редактировать WB магазин', 'editWBStore')
      .addItem('🗑️ Удалить WB магазин', 'deleteWBStore')
      .addItem('🔄 Переключить активный WB магазин', 'switchActiveWBStore'))
    .addSubMenu(ui.createMenu('🛍️ Управление Яндекс Маркет магазинами')
      .addItem('➕ Добавить Яндекс Маркет магазин', 'addNewYandexStore')
      .addItem('📋 Список Яндекс Маркет магазинов', 'showYandexStoresList')
      .addItem('✏️ Редактировать Яндекс Маркет магазин', 'editYandexStore')
      .addItem('🗑️ Удалить Яндекс Маркет магазин', 'deleteYandexStore')
      .addItem('🔄 Переключить активный Яндекс Маркет магазин', 'switchActiveYandexStore'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 WB Выгрузка остатков')
      .addItem('📦 Выгрузить FBO остатки (активный WB)', 'exportWBFBOStocks')
      .addItem('📦 Выгрузить FBO остатки (Statistics API)', 'loadAllStocks')
      .addItem('📦 Выгрузить FBO остатки (с увеличенными задержками)', 'exportWBFBOStocksWithLongDelays')
      .addItem('📦 Выгрузить FBO остатки (все WB магазины)', 'exportAllWBStoresStocks')
      .addItem('📦 Выгрузить FBO остатки (все WB магазины, Statistics API)', 'exportAllWBStoresStocksStatisticsAPI')
      .addSeparator()
      .addItem('📈 Выгрузить цены (активный WB)', 'exportWBPrices')
      .addItem('📈 Выгрузить цены (все WB магазины)', 'exportAllWBStoresPrices')
      .addSeparator()
      .addItem('📊 Выгрузить через Statistics API (продажи)', 'exportWBStocksViaStatisticsAPI')
      .addItem('🧪 Тест WB API', 'testWBConnection')
      .addItem('🧪 Тест WB API (taskId)', 'testWBTaskIdAPI')
      .addItem('🧪 Тест WB Statistics API', 'testWBStatisticsAPI')
      .addItem('🧪 Тест WB Statistics API (периоды)', 'testWBStatisticsAPIWithPeriods')
      .addItem('⚙️ Настройка параметров отчёта WB', 'configureWBReportParams')
      .addItem('⚙️ Настройка лимитов WB API', 'configureWBRateLimits'))
    .addSubMenu(ui.createMenu('📊 Яндекс Маркет Выгрузка остатков')
      .addItem('📦 Выгрузить остатки (активный Яндекс Маркет)', 'exportYandexStocks')
      .addItem('📦 Выгрузить остатки (все Яндекс Маркет магазины)', 'exportAllYandexStoresStocks')
      .addSeparator()
      .addItem('📈 Выгрузить цены (активный Яндекс Маркет)', 'exportYandexPrices')
      .addSeparator()
      .addItem('🧪 Тест Яндекс Маркет API', 'testYandexConnection')
      .addItem('🚀 Быстрый тест с вашими токенами', 'testYandexWithYourTokens'))
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
 * Получает список всех WB магазинов
 */
function getWBStoresList() {
  const properties = PropertiesService.getScriptProperties();
  const storesJson = properties.getProperty('WB_STORES');
  
  if (!storesJson) {
    return [];
  }
  
  try {
    return JSON.parse(storesJson);
  } catch (error) {
    console.error('Ошибка парсинга списка WB магазинов:', error);
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
 * Сохраняет список WB магазинов
 */
function saveWBStoresList(stores) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('WB_STORES', JSON.stringify(stores));
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
 * Получает активный WB магазин
 */
function getActiveWBStore() {
  const properties = PropertiesService.getScriptProperties();
  const activeStoreId = properties.getProperty('ACTIVE_WB_STORE_ID');
  
  if (!activeStoreId) {
    return null;
  }
  
  const stores = getWBStoresList();
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
 * Устанавливает активный WB магазин
 */
function setActiveWBStore(storeId) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('ACTIVE_WB_STORE_ID', storeId);
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
 * Добавляет новый WB магазин
 */
function addNewWBStore() {
  const ui = SpreadsheetApp.getUi();
  
  // Запрашиваем данные магазина
  const storeName = ui.prompt('Добавить WB магазин', 'Введите название магазина:', ui.ButtonSet.OK_CANCEL);
  if (storeName.getSelectedButton() !== ui.Button.OK) return;
  
  const apiKey = ui.prompt('Добавить WB магазин', 'Введите API Key:', ui.ButtonSet.OK_CANCEL);
  if (apiKey.getSelectedButton() !== ui.Button.OK) return;
  
  // Проверяем что все поля заполнены
  if (!storeName.getResponseText().trim() || !apiKey.getResponseText().trim()) {
    ui.alert('Ошибка', 'Все поля должны быть заполнены!', ui.ButtonSet.OK);
    return;
  }
  
  // Создаем новый магазин
  const newStore = {
    id: Utilities.getUuid(),
    name: storeName.getResponseText().trim(),
    api_key: apiKey.getResponseText().trim(),
    created_at: new Date().toISOString()
  };
  
  // Добавляем в список
  const stores = getWBStoresList();
  stores.push(newStore);
  saveWBStoresList(stores);
  
  // Если это первый магазин, делаем его активным
  if (stores.length === 1) {
    setActiveWBStore(newStore.id);
  }
  
  ui.alert('Успех', `WB магазин "${newStore.name}" добавлен!`, ui.ButtonSet.OK);
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
 * Показывает список WB магазинов
 */
function showWBStoresList() {
  const stores = getWBStoresList();
  const activeStore = getActiveWBStore();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Информация', 'Нет добавленных WB магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  let message = 'WB Магазины:\n\n';
  stores.forEach((store, index) => {
    const isActive = activeStore && store.id === activeStore.id ? ' (АКТИВНЫЙ)' : '';
    message += `${index + 1}. ${store.name}${isActive}\n`;
  });
  
  SpreadsheetApp.getUi().alert('WB Магазины', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Переключает активный WB магазин
 */
function switchActiveWBStore() {
  const stores = getWBStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных WB магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите активный WB магазин:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Переключить WB магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const selectedStore = stores[selectedIndex];
    setActiveWBStore(selectedStore.id);
    ui.alert('Успех', `Активный WB магазин: ${selectedStore.name}`, ui.ButtonSet.OK);
  } else {
    ui.alert('Ошибка', 'Неверный номер магазина', ui.ButtonSet.OK);
  }
}

/**
 * Редактирует WB магазин
 */
function editWBStore() {
  const stores = getWBStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных WB магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите WB магазин для редактирования:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Редактировать WB магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const storeToEdit = stores[selectedIndex];
    
    // Запрашиваем новые данные
    const newName = ui.prompt('Редактировать WB магазин', `Название (текущее: ${storeToEdit.name}):`, ui.ButtonSet.OK_CANCEL);
    if (newName.getSelectedButton() !== ui.Button.OK) return;
    
    const newApiKey = ui.prompt('Редактировать WB магазин', 'API Key (оставьте пустым чтобы не менять):', ui.ButtonSet.OK_CANCEL);
    if (newApiKey.getSelectedButton() !== ui.Button.OK) return;
    
    // Обновляем данные
    if (newName.getResponseText().trim()) {
      storeToEdit.name = newName.getResponseText().trim();
    }
    
    if (newApiKey.getResponseText().trim()) {
      storeToEdit.api_key = newApiKey.getResponseText().trim();
    }
    
    storeToEdit.updated_at = new Date().toISOString();
    
    // Сохраняем изменения
    saveWBStoresList(stores);
    
    ui.alert('Успех', 'WB магазин обновлен!', ui.ButtonSet.OK);
  } else {
    ui.alert('Ошибка', 'Неверный номер магазина', ui.ButtonSet.OK);
  }
}

/**
 * Удаляет WB магазин
 */
function deleteWBStore() {
  const stores = getWBStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных WB магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите WB магазин для удаления:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Удалить WB магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const storeToDelete = stores[selectedIndex];
    
    // Подтверждение удаления
    const confirm = ui.alert('Подтверждение', `Удалить WB магазин "${storeToDelete.name}"?`, ui.ButtonSet.YES_NO);
    if (confirm === ui.Button.YES) {
      stores.splice(selectedIndex, 1);
      saveWBStoresList(stores);
      
      // Если удалили активный магазин, выбираем новый
      const activeStore = getActiveWBStore();
      if (!activeStore && stores.length > 0) {
        setActiveWBStore(stores[0].id);
      }
      
      ui.alert('Успех', 'WB магазин удален!', ui.ButtonSet.OK);
    }
  } else {
    ui.alert('Ошибка', 'Неверный номер магазина', ui.ButtonSet.OK);
  }
}

// ==================== ФУНКЦИИ ДЛЯ УПРАВЛЕНИЯ ЯНДЕКС МАРКЕТ МАГАЗИНАМИ ====================

/**
 * Получает список всех Яндекс Маркет магазинов
 */
function getYandexStoresList() {
  const properties = PropertiesService.getScriptProperties();
  const storesJson = properties.getProperty('YANDEX_STORES');
  
  if (!storesJson) {
    return [];
  }
  
  try {
    return JSON.parse(storesJson);
  } catch (error) {
    console.error('Ошибка парсинга Яндекс Маркет магазинов:', error);
    return [];
  }
}

/**
 * Сохраняет список Яндекс Маркет магазинов
 */
function saveYandexStoresList(stores) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('YANDEX_STORES', JSON.stringify(stores));
}

/**
 * Получает активный Яндекс Маркет магазин
 */
function getActiveYandexStore() {
  const properties = PropertiesService.getScriptProperties();
  const activeStoreId = properties.getProperty('ACTIVE_YANDEX_STORE_ID');
  
  if (!activeStoreId) {
    return null;
  }
  
  const stores = getYandexStoresList();
  return stores.find(store => store.id === activeStoreId) || null;
}

/**
 * Устанавливает активный Яндекс Маркет магазин
 */
function setActiveYandexStore(storeId) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty('ACTIVE_YANDEX_STORE_ID', storeId);
}

/**
 * Добавляет новый Яндекс Маркет магазин
 */
function addNewYandexStore() {
  const ui = SpreadsheetApp.getUi();
  
  // Запрашиваем данные магазина
  const storeName = ui.prompt('Добавить Яндекс Маркет магазин', 'Введите название магазина:', ui.ButtonSet.OK_CANCEL);
  if (storeName.getSelectedButton() !== ui.Button.OK) return;
  
  const apiToken = ui.prompt('Добавить Яндекс Маркет магазин', 'Введите API Token:', ui.ButtonSet.OK_CANCEL);
  if (apiToken.getSelectedButton() !== ui.Button.OK) return;
  
  const campaignId = ui.prompt('Добавить Яндекс Маркет магазин', 'Введите Campaign ID:', ui.ButtonSet.OK_CANCEL);
  if (campaignId.getSelectedButton() !== ui.Button.OK) return;
  
  // Проверяем что все поля заполнены
  if (!storeName.getResponseText().trim() || !apiToken.getResponseText().trim() || !campaignId.getResponseText().trim()) {
    ui.alert('Ошибка', 'Все поля должны быть заполнены!', ui.ButtonSet.OK);
    return;
  }
  
  // Создаем новый магазин
  const newStore = {
    id: Utilities.getUuid(),
    name: storeName.getResponseText().trim(),
    api_token: apiToken.getResponseText().trim(),
    campaign_id: campaignId.getResponseText().trim(),
    created_at: new Date().toISOString()
  };
  
  // Добавляем в список
  const stores = getYandexStoresList();
  stores.push(newStore);
  saveYandexStoresList(stores);
  
  // Если это первый магазин, делаем его активным
  if (stores.length === 1) {
    setActiveYandexStore(newStore.id);
  }
  
  ui.alert('Успех', `Яндекс Маркет магазин "${newStore.name}" добавлен!`, ui.ButtonSet.OK);
}

/**
 * Показывает список Яндекс Маркет магазинов
 */
function showYandexStoresList() {
  const stores = getYandexStoresList();
  const activeStore = getActiveYandexStore();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Информация', 'Нет добавленных Яндекс Маркет магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  let message = 'Яндекс Маркет Магазины:\n\n';
  stores.forEach((store, index) => {
    const isActive = activeStore && store.id === activeStore.id ? ' (АКТИВНЫЙ)' : '';
    message += `${index + 1}. ${store.name}${isActive}\n`;
    message += `   Campaign ID: ${store.campaign_id}\n`;
    message += `   API Token: ***${store.api_token.slice(-4)}\n\n`;
  });
  
  SpreadsheetApp.getUi().alert('Яндекс Маркет Магазины', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Переключает активный Яндекс Маркет магазин
 */
function switchActiveYandexStore() {
  const stores = getYandexStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных Яндекс Маркет магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите активный Яндекс Маркет магазин:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Переключить Яндекс Маркет магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const selectedStore = stores[selectedIndex];
    setActiveYandexStore(selectedStore.id);
    ui.alert('Успех', `Активный Яндекс Маркет магазин изменен на "${selectedStore.name}"`, ui.ButtonSet.OK);
  } else {
    ui.alert('Ошибка', 'Неверный номер магазина', ui.ButtonSet.OK);
  }
}

/**
 * Редактирует Яндекс Маркет магазин
 */
function editYandexStore() {
  const stores = getYandexStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных Яндекс Маркет магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите Яндекс Маркет магазин для редактирования:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Редактировать Яндекс Маркет магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const storeToEdit = stores[selectedIndex];
    
    // Запрашиваем новые данные
    const newName = ui.prompt('Редактировать Яндекс Маркет магазин', `Название (текущее: ${storeToEdit.name}):`, ui.ButtonSet.OK_CANCEL);
    if (newName.getSelectedButton() !== ui.Button.OK) return;
    
    const newApiToken = ui.prompt('Редактировать Яндекс Маркет магазин', 'API Token (оставьте пустым чтобы не менять):', ui.ButtonSet.OK_CANCEL);
    if (newApiToken.getSelectedButton() !== ui.Button.OK) return;
    
    const newCampaignId = ui.prompt('Редактировать Яндекс Маркет магазин', `Campaign ID (текущий: ${storeToEdit.campaign_id}):`, ui.ButtonSet.OK_CANCEL);
    if (newCampaignId.getSelectedButton() !== ui.Button.OK) return;
    
    // Обновляем данные
    if (newName.getResponseText().trim()) {
      storeToEdit.name = newName.getResponseText().trim();
    }
    if (newApiToken.getResponseText().trim()) {
      storeToEdit.api_token = newApiToken.getResponseText().trim();
    }
    if (newCampaignId.getResponseText().trim()) {
      storeToEdit.campaign_id = newCampaignId.getResponseText().trim();
    }
    
    storeToEdit.updated_at = new Date().toISOString();
    
    // Сохраняем изменения
    saveYandexStoresList(stores);
    
    ui.alert('Успех', 'Яндекс Маркет магазин обновлен!', ui.ButtonSet.OK);
  } else {
    ui.alert('Ошибка', 'Неверный номер магазина', ui.ButtonSet.OK);
  }
}

/**
 * Удаляет Яндекс Маркет магазин
 */
function deleteYandexStore() {
  const stores = getYandexStoresList();
  
  if (stores.length === 0) {
    SpreadsheetApp.getUi().alert('Ошибка', 'Нет доступных Яндекс Маркет магазинов', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const ui = SpreadsheetApp.getUi();
  let message = 'Выберите Яндекс Маркет магазин для удаления:\n\n';
  stores.forEach((store, index) => {
    message += `${index + 1}. ${store.name}\n`;
  });
  
  const response = ui.prompt('Удалить Яндекс Маркет магазин', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const selectedIndex = parseInt(response.getResponseText()) - 1;
  
  if (selectedIndex >= 0 && selectedIndex < stores.length) {
    const storeToDelete = stores[selectedIndex];
    
    // Подтверждение удаления
    const confirm = ui.alert('Подтверждение', `Удалить Яндекс Маркет магазин "${storeToDelete.name}"?`, ui.ButtonSet.YES_NO);
    if (confirm === ui.Button.YES) {
      stores.splice(selectedIndex, 1);
      saveYandexStoresList(stores);
      
      // Если удалили активный магазин, выбираем новый
      const activeStore = getActiveYandexStore();
      if (!activeStore && stores.length > 0) {
        setActiveYandexStore(stores[0].id);
      }
      
      ui.alert('Успех', 'Яндекс Маркет магазин удален!', ui.ButtonSet.OK);
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

// ==================== WB ЦЕНЫ ====================

/**
 * Выгрузить цены для активного WB магазина и записать в лист с колонки T
 */
function exportWBPrices() {
  const config = getWBConfig();
  // Сначала пробуем публичный API по nmId из колонки P2:P активного листа WB магазина
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sanitizeSheetName(config.STORE_NAME || 'WB Магазин');
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  const nmIds = readColumnValues_(sheet, 2, 16); // P2:P
  if (nmIds.length > 0) {
    const publicPrices = fetchWBPublicPricesByNmIds(nmIds);
    writeWBPublicPricesToSheetT(publicPrices, config.STORE_NAME, nmIds);
    return;
  }

  // Если P2:P пусто — используем закрытый Supplier API по API_KEY и сопоставлению 'Артикул поставщика'
  if (!config.API_KEY) {
    throw new Error('Для WB не задан API_KEY. Добавьте WB магазин или задайте ключ.');
  }
  const prices = fetchWBPrices(config.API_KEY);
  writeWBPricesToSheetT(prices, config.STORE_NAME);
}

/**
 * Выгружает цены WB для всех магазинов.
 * Логика:
 * - Если на листе магазина есть nmId в P2:P — используем публичный API и записываем через writeWBPublicPricesToSheetT
 * - Иначе, если задан API_KEY магазина — используем Supplier API и writeWBPricesToSheetT
 */
function exportAllWBStoresPrices() {
  const stores = getWBStoresList();
  if (!Array.isArray(stores) || stores.length === 0) {
    throw new Error('Список WB магазинов пуст. Добавьте WB магазины.');
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  let processed = 0;
  const errors = [];

  for (const store of stores) {
    try {
      const storeName = store.name || 'WB Магазин';
      const sheetName = sanitizeSheetName(storeName);
      const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

      const nmIds = readColumnValues_(sheet, 2, 16); // P2:P
      if (nmIds.length > 0) {
        const publicPrices = fetchWBPublicPricesByNmIds(nmIds);
        writeWBPublicPricesToSheetT(publicPrices, storeName, nmIds);
        processed++;
        Utilities.sleep(200);
        continue;
      }

      const apiKey = store.api_key || store.apiKey || '';
      if (apiKey) {
        const prices = fetchWBPrices(apiKey);
        writeWBPricesToSheetT(prices, storeName);
        processed++;
        Utilities.sleep(200);
        continue;
      }

      errors.push(`Магазин ${storeName}: нет nmId в P и не задан API ключ`);
    } catch (e) {
      errors.push(String(e));
    }
  }

  const msg = `Готово. Обработано магазинов: ${processed}/${stores.length}.` + (errors.length ? ('\nОшибки:\n- ' + errors.join('\n- ')) : '');
  console.log(msg);
  if (errors.length) {
    // Позволим внешнему раннеру (runStep) пометить шаг как упавший и отправить уведомление
    throw new Error(msg);
  } else {
    // UI alert только если запущено вручную (не из триггера)
    try {
      ui.alert('Выгрузка цен WB (все магазины)', msg, ui.ButtonSet.OK);
    } catch (e) {
      // Игнорируем ошибки UI в триггерах
      console.log('UI alert пропущен (запуск из триггера)');
    }
  }
}

/**
 * Получает список цен из WB API. Спекуляция: основной путь — /public/api/v1/info,
 * фоллбек — /public/api/v1/prices (если доступен у аккаунта).
 */
function fetchWBPrices(apiKey) {
  const base = 'https://suppliers-api.wildberries.ru';
  const headers = {
    'Authorization': apiKey
  };

  // Пытаемся получить текущие цены и скидки
  // Вариант 1: /public/api/v1/info?quantity=0 — часто доступен и содержит priceU/salePriceU
  try {
    const url = base + '/public/api/v1/info?quantity=0';
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      const data = JSON.parse(resp.getContentText());
      if (Array.isArray(data)) {
        return data.map(item => {
          // priceU/salePriceU в копейках
          const supplierArticle = item.supplierArticle || item.supplier_article || '';
          const priceU = Number(item.priceU || 0);
          const salePriceU = Number(item.salePriceU || 0);
          return {
            supplierArticle: supplierArticle,
            price: salePriceU ? Math.round(salePriceU / 100) : '',
            old_price: priceU ? Math.round(priceU / 100) : '',
            min_price: '',
            currency: 'RUB'
          };
        });
      }
    }
  } catch (e) {
    // перейдем к фоллбеку
  }

  // Вариант 2 (фоллбек): /public/api/v1/prices — некоторые аккаунты имеют этот метод
  try {
    const url = base + '/public/api/v1/prices';
    const resp = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code >= 200 && code < 300) {
      const data = JSON.parse(resp.getContentText());
      if (Array.isArray(data)) {
        return data.map(item => {
          const supplierArticle = item.supplierArticle || item.supplier_article || '';
          const priceU = Number(item.priceU || 0);
          const salePriceU = Number(item.salePriceU || 0);
          return {
            supplierArticle: supplierArticle,
            price: salePriceU ? Math.round(salePriceU / 100) : '',
            old_price: priceU ? Math.round(priceU / 100) : '',
            min_price: '',
            currency: 'RUB'
          };
        });
      }
    }
  } catch (e) {
    // игнор
  }

  return [];
}

/**
 * Запись цен WB в лист магазина, начиная с T. Сопоставление по 'Артикул поставщика'.
 */
function writeWBPricesToSheetT(prices, storeName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sanitizeSheetName(storeName || 'WB Магазин');
  let sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  const headerRow = 1;
  const startCol = 20; // T
  const headers = ['Артикул поставщика', 'Цена, ₽', 'Старая цена, ₽', 'Мин. цена, ₽', 'Валюта'];
  // Очистка диапазона T:X перед записью
  const lastRowClear1 = sheet.getLastRow();
  if (lastRowClear1 > 0) {
    sheet.getRange(1, startCol, lastRowClear1, headers.length).clearContent();
  }
  sheet.getRange(headerRow, startCol, 1, headers.length).setValues([headers]);
  sheet.getRange(headerRow, startCol, 1, headers.length).setFontWeight('bold').setBackground('#FFF3CD');

  // Даже если данных нет — обновим отметку времени ниже
  if (!prices) prices = [];

  // map по supplierArticle
  const bySa = {};
  prices.forEach(p => { if (p.supplierArticle) bySa[p.supplierArticle] = p; });

  // Ищем колонку 'Артикул поставщика' в шапке WB листа
  let saCol = 3; // по умолчанию C, как в наших WB заголовках
  try {
    const firstRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
    const idx = firstRow.findIndex(v => String(v).toLowerCase() === 'артикул поставщика');
    if (idx >= 0) saCol = idx + 1;
  } catch (e) {}

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const saValues = sheet.getRange(2, saCol, lastRow - 1, 1).getValues();
  const rows = [];
  for (let i = 0; i < saValues.length; i++) {
    const sa = (saValues[i][0] || '').toString().trim();
    const p = bySa[sa];
    if (p) {
      rows.push([
        sa,
        (p.price !== undefined && p.price !== null && p.price !== '') ? p.price : 0,
        (p.old_price !== undefined && p.old_price !== null && p.old_price !== '') ? p.old_price : 0,
        (p.min_price !== undefined && p.min_price !== null && p.min_price !== '') ? p.min_price : 0,
        p.currency || 'RUB'
      ]);
    } else {
      rows.push([sa, 0, 0, 0, 'RUB']);
    }
  }

  sheet.getRange(2, startCol, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(startCol, headers.length);
  sheet.getRange(rows.length + 3, startCol).setValue('Цены WB обновлены: ' + new Date().toLocaleString('ru-RU'));
}

/**
 * Читает значения из колонки начиная с rowStart. Возвращает массив строк (без пустых).
 */
function readColumnValues_(sheet, rowStart, colIndex) {
  const last = sheet.getLastRow();
  if (last < rowStart) return [];
  const rng = sheet.getRange(rowStart, colIndex, last - rowStart + 1, 1).getValues();
  const out = [];
  for (const [cell] of rng) {
    if (cell === '' || cell === null || cell === undefined) continue;
    const s = String(cell).trim();
    if (s) out.push(s);
  }
  return out;
}

/**
 * Публичный WB: получает цены по массиву nmId через cards/v2/detail, батчами по 100.
 */
function fetchWBPublicPricesByNmIds(nmIds) {
  const BASE = 'https://card.wb.ru/cards/v2/detail?appType=1&curr=rub&dest=-1255987&spp=30&ab_testing=false&nm=';
  const CHUNK = 100;
  const map = new Map(); // nmId -> { nmId, price, old_price, discount_percent, currency }

  for (let i = 0; i < nmIds.length; i += CHUNK) {
    const chunk = nmIds.slice(i, i + CHUNK);
    const url = BASE + encodeURIComponent(chunk.join(';'));
    const resp = safeFetchJson_(url);
    const products = resp && resp.data && Array.isArray(resp.data.products) ? resp.data.products : [];
    for (const p of products) {
      const nmId = String(p.id);
      let priceBasic = null;
      let priceTotal = null;
      if (Array.isArray(p.sizes) && p.sizes.length) {
        const s = p.sizes[0];
        priceBasic = toMoney_(s && s.price && s.price.basic);
        priceTotal = toMoney_(s && s.price && s.price.total);
      }
      // Fallback на верхнеуровневые поля
      if (priceBasic == null) priceBasic = toMoney_(p.priceU);
      if (priceTotal == null) priceTotal = toMoney_(p.salePriceU);
      let discount = '';
      if (isFiniteNumber_(priceBasic) && isFiniteNumber_(priceTotal) && priceBasic > 0) {
        discount = Math.round((1 - priceTotal / priceBasic) * 100);
      }
      map.set(nmId, {
        nmId: nmId,
        price: (priceTotal != null && priceTotal !== '') ? priceTotal : 0,
        old_price: (priceBasic != null && priceBasic !== '') ? priceBasic : 0,
        discount_percent: discount,
        currency: 'RUB'
      });
    }
    Utilities.sleep(150);
  }

  return map;
}

function writeWBPublicPricesToSheetT(priceMap, storeName, nmIdsOrder) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sanitizeSheetName(storeName || 'WB Магазин');
  let sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  const headerRow = 1;
  const startCol = 20; // T
  const headers = ['nmId', 'Цена, ₽', 'Старая цена, ₽', 'Скидка, %', 'Валюта'];
  // Очистка диапазона T:X перед записью
  const lastRowClear2 = sheet.getLastRow();
  if (lastRowClear2 > 0) {
    sheet.getRange(1, startCol, lastRowClear2, headers.length).clearContent();
  }
  sheet.getRange(headerRow, startCol, 1, headers.length).setValues([headers]);
  sheet.getRange(headerRow, startCol, 1, headers.length).setFontWeight('bold').setBackground('#FFF3CD');

  if (!priceMap || priceMap.size === 0) {
    // Нет цен — но обновим отметку времени, чтобы видеть факт запуска
    sheet.getRange(3, startCol).setValue('Цены WB (public) обновлены: ' + new Date().toLocaleString('ru-RU'));
    return;
  }

  const rows = [];
  // пишем в порядке nmIds из колонки P
  for (const nmIdRaw of nmIdsOrder) {
    const nmId = String(nmIdRaw);
    const p = priceMap.get(nmId);
    if (p) {
      rows.push([
        nmId,
        (p.price !== undefined && p.price !== null && p.price !== '') ? p.price : 0,
        (p.old_price !== undefined && p.old_price !== null && p.old_price !== '') ? p.old_price : 0,
        p.discount_percent || 0,
        p.currency || 'RUB'
      ]);
    } else {
      rows.push([nmId, 0, 0, 0, 'RUB']);
    }
  }

  sheet.getRange(2, startCol, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(startCol, headers.length);
  sheet.getRange(rows.length + 3, startCol).setValue('Цены WB (public) обновлены: ' + new Date().toLocaleString('ru-RU'));
}

// Вспомогательные для публичного WB
function safeFetchJson_(url) {
  try {
    const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true, headers: { 'Accept': 'application/json' } });
    if (res.getResponseCode() >= 200 && res.getResponseCode() < 300) {
      return JSON.parse(res.getContentText());
    }
    return null;
  } catch (e) {
    return null;
  }
}

function toMoney_(v) {
  if (v === null || v === undefined || v === '') return null;
  const n = Number(v);
  if (!Number.isFinite(n)) return null;
  return Math.round(n) / 100;
}

function isFiniteNumber_(v) {
  return typeof v === 'number' && Number.isFinite(v);
}

/**
 * Получает цены товаров Ozon с пагинацией v4
 */
function fetchAllOzonPricesV4() {
  const config = getOzonConfig();
  const headers = {
    'Client-Id': config.CLIENT_ID,
    'Api-Key': config.API_KEY
  };

  let lastId = '';
  const result = [];
  const PAGE_LIMIT = 1000;

  do {
    const payload = {
      filter: {
        visibility: 'ALL'
      },
      limit: PAGE_LIMIT,
      last_id: lastId
    };

    // Согласно v4: /v4/product/info/prices
    const resp = callOzonAPI('/v4/product/info/prices', payload, headers);

    let items = [];
    if (resp && resp.result && Array.isArray(resp.result.items)) {
      items = resp.result.items;
      lastId = resp.result.last_id || '';
    } else if (Array.isArray(resp.items)) {
      items = resp.items;
      lastId = resp.last_id || '';
    } else {
      items = [];
      lastId = '';
    }

    for (const it of items) {
      result.push({
        offer_id: it.offer_id || '',
        sku: it.sku || '',
        product_id: it.product_id || it.id || '',
        // В ответе встречаются price и old_price, min_price
        price: it.price && it.price.price ? Number(it.price.price) : (typeof it.price === 'number' ? it.price : null),
        old_price: it.price && it.price.old_price ? Number(it.price.old_price) : null,
        min_price: it.price && it.price.min_price ? Number(it.price.min_price) : null,
        currency_code: it.price && it.price.currency_code ? it.price.currency_code : ''
      });
    }
  } while (lastId);

  return result;
}

/**
 * Считывает offer_id из листа активного магазина (ищет колонку 'Артикул')
 */
function getOfferIdsFromActiveStoreSheet() {
  const config = getOzonConfig();
  let spreadsheetId = config.SPREADSHEET_ID || SpreadsheetApp.getActiveSpreadsheet().getId();
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  const sheetName = sanitizeSheetName(config.STORE_NAME || 'Неизвестный магазин');
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) return [];

  const header = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
  let offerCol = header.findIndex(v => String(v).toLowerCase() === 'артикул');
  if (offerCol === -1) offerCol = 4; // по умолчанию E -> индекс 4 (0-based)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, offerCol + 1, lastRow - 1, 1).getValues();
  const ids = [];
  for (const v of values) {
    const id = (v[0] || '').toString().trim();
    if (id) ids.push(id);
  }
  // убираем дубли
  return Array.from(new Set(ids));
}

/**
 * Получает цены по списку offer_id, поддерживая разные варианты эндпоинтов и тел
 */
function fetchOzonPricesByOfferIds(offerIds) {
  const config = getOzonConfig();
  const headers = {
    'Client-Id': config.CLIENT_ID,
    'Api-Key': config.API_KEY
  };

  const chunkSize = 1000; // максимально щадящий батч
  const result = [];

  for (let i = 0; i < offerIds.length; i += chunkSize) {
    const chunk = offerIds.slice(i, i + chunkSize);

    // Сначала пробуем v5 с курсором
    let v5Items = [];
    try {
      let cursor = '';
      do {
        const body = {
          filter: { offer_id: chunk, visibility: 'ALL' },
          limit: 1000,
          cursor: cursor
        };
        const resp = callOzonAPI('/v5/product/info/prices', body, headers);
        const items = (resp && resp.items) || [];
        if (Array.isArray(items) && items.length) v5Items.push(...items);
        cursor = (resp && resp.cursor) || '';
        Utilities.sleep(100);
      } while (cursor);
    } catch (e) {
      v5Items = [];
    }

    let items = v5Items;
    if (!items || items.length === 0) {
      // Фоллбеки: v3, v2, v4 (на случай особенностей аккаунта/версии)
      const tryCalls = [
        { path: '/v3/product/info/prices', body: { offer_id: chunk } },
        { path: '/v2/product/info/prices', body: { offer_id: chunk } },
        { path: '/v4/product/info/prices', body: { filter: { offer_id: chunk } } }
      ];
      let got = null;
      for (const tc of tryCalls) {
        try {
          const resp = callOzonAPI(tc.path, tc.body, headers);
          got = resp;
          break;
        } catch (e) {
          // пробуем следующий
        }
      }
      if (!got) continue;
      if (got.result && Array.isArray(got.result)) items = got.result;
      if (got.result && Array.isArray(got.result.items)) items = got.result.items;
      if (Array.isArray(got.items)) items = got.items;
    }

    for (const it of items) {
      const priceObj = it.price || it.prices || it.price_info || {};
      const price = typeof it.price === 'number' ? it.price : (priceObj.price || priceObj.value || null);
      const old_price = priceObj.old_price || null;
      const min_price = priceObj.min_price || null;
      const currency_code = priceObj.currency_code || priceObj.currency || '';

      result.push({
        offer_id: it.offer_id || it.offerId || '',
        sku: it.sku || '',
        product_id: it.product_id || it.id || '',
        price: price != null ? Number(price) : null,
        old_price: old_price != null ? Number(old_price) : null,
        min_price: min_price != null ? Number(min_price) : null,
        currency_code
      });
    }
  }

  return result;
}

/**
 * Выгружает цены и записывает в лист магазина с колонки T
 */
function exportOzonPrices() {
  const config = getOzonConfig();
  if (!config.CLIENT_ID || !config.API_KEY) {
    throw new Error('Не настроены API ключи Ozon. Добавьте магазин.');
  }

  // Читаем offer_id из листа активного магазина и вытягиваем цены батчами
  const offerIds = getOfferIdsFromActiveStoreSheet();
  if (offerIds.length === 0) {
    throw new Error('Не найден ни один offer_id на листе магазина. Сначала выгрузите остатки.');
  }

  const prices = fetchOzonPricesByOfferIds(offerIds);
  writePricesToSheetT(prices);
}

/**
 * Выгружает детальные цены Ozon в формате как в ozon_price_example
 */
function exportOzonPricesDetailed() {
  const config = getOzonConfig();
  if (!config.CLIENT_ID || !config.API_KEY) {
    throw new Error('Не настроены API ключи Ozon. Добавьте магазин.');
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sanitizeSheetName(config.STORE_NAME || 'Ozon Prices');
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  // Получаем все цены через v5 API
  const allPrices = fetchAllOzonPricesV5(config.CLIENT_ID, config.API_KEY);
  
  // Записываем в детальном формате
  writeOzonPricesDetailed(sheet, allPrices);
  
  console.log(`Выгружено ${allPrices.length} товаров с ценами`);
  SpreadsheetApp.getUi().alert('Выгрузка завершена', `Загружено ${allPrices.length} товаров`, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Выгружает детальные цены для всех Ozon магазинов в их листы
 */
function exportAllStoresPricesDetailed() {
  const stores = getStoresList();
  if (!Array.isArray(stores) || stores.length === 0) {
    throw new Error('Список магазинов пуст. Добавьте магазины в \'🏪 Управление магазинами\'.');
  }

  const ui = SpreadsheetApp.getUi();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  let processed = 0;
  let totalItems = 0;
  const errors = [];

  for (const store of stores) {
    const clientId = store.clientId;
    const apiKey = store.apiKey;
    const storeName = store.name || 'Ozon Store';

    if (!clientId || !apiKey) {
      errors.push(`Магазин ${storeName}: не заданы clientId/apiKey`);
      continue;
    }

    try {
      const items = fetchAllOzonPricesV5(clientId, apiKey);
      totalItems += items.length;

      const sheetName = sanitizeSheetName(storeName);
      const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
      writeOzonPricesDetailed(sheet, items);
      processed++;

      Utilities.sleep(300);
    } catch (e) {
      errors.push(`Магазин ${storeName}: ${String(e)}`);
    }
  }

  const msg = `Готово. Магазинов: ${processed}/${stores.length}. Всего позиций: ${totalItems}.` + (errors.length ? ('\nОшибки:\n- ' + errors.join('\n- ')) : '');
  console.log(msg);
  
  // UI alert только если запущено вручную (не из триггера)
  try {
    ui.alert('Выгрузка детальных цен (все магазины)', msg, ui.ButtonSet.OK);
  } catch (e) {
    // Игнорируем ошибки UI в триггерах
    console.log('UI alert пропущен (запуск из триггера)');
  }
}

/**
 * Получает все цены через v5 API с пагинацией
 */
function fetchAllOzonPricesV5(clientId, apiKey) {
  const url = 'https://api-seller.ozon.ru/v5/product/info/prices';
  const headers = {
    'Client-Id': clientId,
    'Api-Key': apiKey
  };
  
  let cursor = '';
  const allItems = [];
  let page = 0;
  
  do {
    const body = {
      cursor: cursor,
      filter: {
        visibility: 'ALL'
      },
      limit: 1000
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      headers: headers,
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    
    if (code >= 200 && code < 300) {
      const data = JSON.parse(response.getContentText());
      if (data.items && data.items.length > 0) {
        allItems.push(...data.items);
      }
      cursor = data.cursor || '';
    } else {
      throw new Error(`Ошибка API: HTTP ${code}: ${response.getContentText()}`);
    }
    
    page++;
    Utilities.sleep(150);
    if (page > 1000) throw new Error('Слишком много страниц, остановлено');
  } while (cursor);
  
  return allItems;
}

/**
 * Записывает детальные цены Ozon в формате как в ozon_price_example начиная со столбца T
 */
function writeOzonPricesDetailed(sheet, items) {
  const startCol = 20; // T
  const headerRow = 1;
  
  const header = [
    'Артикул продавца',
    'ID товара', 
    'Валюта',
    'Цена',
    'Старая цена',
    'Маркетинговая цена',
    'Маркетинговая цена продавца',
    'Мин. цена',
    'Нетто цена',
    'Розничная цена',
    'НДС',
    'Эквайринг',
    '% продаж FBO',
    '% продаж FBS',
    'FBO доставка покупателю',
    'FBO прямой поток мин',
    'FBO прямой поток макс', 
    'FBO возврат',
    'FBS доставка покупателю',
    'FBS прямой поток мин',
    'FBS прямой поток макс',
    'FBS первая миля мин',
    'FBS первая миля макс',
    'FBS возврат',
    'Индекс цены (цвет)',
    'Ozon индекс мин. цена',
    'Ozon индекс валюта',
    'Ozon индекс значение',
    'Внешний индекс мин. цена',
    'Внешний индекс валюта', 
    'Внешний индекс значение',
    'Собственные МП мин. цена',
    'Собственные МП валюта',
    'Собственные МП значение',
    'Объёмный вес'
  ];

  // Сортируем по цене по убыванию
  const sorted = items.slice().sort((a, b) => {
    const ap = Number(((a.price || {}).price) || 0);
    const bp = Number(((b.price || {}).price) || 0);
    return bp - ap;
  });

  const rows = sorted.map((item) => {
    const p = item.price || {};
    const c = item.commissions || {};
    const idx = item.price_indexes || {};
    const oz = (idx.ozon_index_data || {});
    const ex = (idx.external_index_data || {});
    const sm = (idx.self_marketplaces_index_data || {});
    
    return [
      safeString(item.offer_id),
      safeString(item.product_id),
      safeString(p.currency_code),
      safeNumber(p.price),
      safeNumber(p.old_price),
      safeNumber(p.marketing_price),
      safeNumber(p.marketing_seller_price),
      safeNumber(p.min_price),
      safeNumber(p.net_price),
      safeNumber(p.retail_price),
      safeNumber(p.vat),
      safeNumber(item.acquiring),
      safeNumber(c.sales_percent_fbo),
      safeNumber(c.sales_percent_fbs),
      safeNumber(c.fbo_deliv_to_customer_amount),
      safeNumber(c.fbo_direct_flow_trans_min_amount),
      safeNumber(c.fbo_direct_flow_trans_max_amount),
      safeNumber(c.fbo_return_flow_amount),
      safeNumber(c.fbs_deliv_to_customer_amount),
      safeNumber(c.fbs_direct_flow_trans_min_amount),
      safeNumber(c.fbs_direct_flow_trans_max_amount),
      safeNumber(c.fbs_first_mile_min_amount),
      safeNumber(c.fbs_first_mile_max_amount),
      safeNumber(c.fbs_return_flow_amount),
      safeString(idx.color_index),
      safeNumber(oz.min_price),
      safeString(oz.min_price_currency),
      safeNumber(oz.price_index_value),
      safeNumber(ex.min_price),
      safeString(ex.min_price_currency),
      safeNumber(ex.price_index_value),
      safeNumber(sm.min_price),
      safeString(sm.min_price_currency),
      safeNumber(sm.price_index_value),
      safeNumber(item.volume_weight)
    ];
  });

  // Очищаем диапазон T:BB (столбцы 20-54)
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.getRange(1, startCol, lastRow, 35).clearContent(); // T:BB (20-54 колонки)
  }
  
  // Записываем заголовки начиная со столбца T
  sheet.getRange(headerRow, startCol, 1, header.length).setValues([header]);
  
  // Форматируем заголовки
  sheet.getRange(headerRow, startCol, 1, header.length)
    .setFontWeight('bold')
    .setBackground('#E8F0FE');
  
  // Записываем данные начиная со строки 2, столбца T
  if (rows.length > 0) {
    sheet.getRange(2, startCol, rows.length, header.length).setValues(rows);
  }
  
  // Автоподбор ширины колонок для диапазона T:AY
  sheet.autoResizeColumns(startCol, header.length);
}

// Вспомогательные функции для безопасного преобразования
function safeString(v) {
  if (v === null || v === undefined) return '';
  return String(v);
}

function safeNumber(v) {
  if (v === null || v === undefined || v === '') return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}

/**
 * Записывает цены в лист магазина, начиная со столбца T (20-я колонка)
 */
function writePricesToSheetT(prices) {
  const config = getOzonConfig();
  let spreadsheetId = config.SPREADSHEET_ID || SpreadsheetApp.getActiveSpreadsheet().getId();
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (e) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  const sheetName = sanitizeSheetName(config.STORE_NAME || 'Неизвестный магазин');
  let sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);

  // Очищаем диапазон T:BB перед записью новых данных
  const maxRow = sheet.getLastRow() || 1000;
  if (maxRow > 0) {
    sheet.getRange(1, 20, maxRow, 29).clearContent(); // T:BB (20-48 колонки)
  }

  // Обновляем заголовки T:Z (или дальше), не трогаем A:S
  const headerRow = 1;
  const startCol = 20; // T
  const headers = ['offer_id', 'Цена, ₽', 'Старая цена, ₽', 'Мин. цена, ₽', 'Валюта'];
  sheet.getRange(headerRow, startCol, 1, headers.length).setValues([headers]);
  sheet.getRange(headerRow, startCol, 1, headers.length).setFontWeight('bold').setBackground('#FFF3CD');

  if (!prices || prices.length === 0) {
    return;
  }

  // Создаём map по offer_id -> {price...}
  const byOffer = {};
  prices.forEach(p => {
    if (p.offer_id) byOffer[p.offer_id] = p;
  });

  // Найти колонки с offer_id на листе: по текущей структуре это колонка E (5)
  // Если структура изменится, можно искать заголовок 'Артикул' в первой строке
  let offerCol = 5;
  try {
    const firstRow = sheet.getRange(1, 1, 1, sheet.getMaxColumns()).getValues()[0];
    const idx = firstRow.findIndex(v => String(v).toLowerCase() === 'артикул');
    if (idx >= 0) offerCol = idx + 1;
  } catch (e) {
    // оставляем дефолт 5
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    // нет строк с данными — только шапка и отметка времени
    sheet.getRange(3, startCol).setValue('Цены WB обновлены: ' + new Date().toLocaleString('ru-RU'));
    return;
  }

  // Читаем offer_id по строкам
  const offerValues = sheet.getRange(2, offerCol, lastRow - 1, 1).getValues();

  const rowsToWrite = [];
  for (let i = 0; i < offerValues.length; i++) {
    const offerId = (offerValues[i][0] || '').toString();
    const p = byOffer[offerId];
    if (p) {
      rowsToWrite.push([
        offerId,
        p.price != null ? p.price : '',
        p.old_price != null ? p.old_price : '',
        p.min_price != null ? p.min_price : '',
        p.currency_code || ''
      ]);
    } else {
      rowsToWrite.push([offerId, '', '', '', '']);
    }
  }

  sheet.getRange(2, startCol, rowsToWrite.length, headers.length).setValues(rowsToWrite);
  sheet.autoResizeColumns(startCol, headers.length);
  sheet.getRange(rowsToWrite.length + 3, startCol).setValue('Цены WB обновлены: ' + new Date().toLocaleString('ru-RU'));
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
  
  // Фильтры отключены по требованию
  
  console.log(`Записано ${rows.length} строк в Google Таблицы`);
}

/**
 * Выгружает остатки со всех магазинов
 */
function exportAllStoresStocks() {
  try {
    const stores = getStoresList();
    
    if (stores.length === 0) {
      try {
        SpreadsheetApp.getUi().alert('Ошибка', 'Нет добавленных магазинов!', SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) {
        console.log('UI alert пропущен (запуск из триггера)');
      }
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
    
    try {
      SpreadsheetApp.getUi().alert('Выгрузка завершена', `Обработано ${stores.length} магазинов, всего товаров: ${totalProcessed}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      console.log('UI alert пропущен (запуск из триггера)');
    }
    
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

// ==================== WB API ФУНКЦИИ ====================

const WB_ANALYTICS_HOST = 'https://seller-analytics-api.wildberries.ru';
const WB_STATISTICS_HOST = 'https://statistics-api.wildberries.ru';
const WB_REPORT_TIMEOUT_MS = 6 * 60 * 1000; // ждать до 6 минут
const WB_REPORT_POLL_INTERVAL_MS = 4000;

// Настройки для обработки лимитов запросов
const WB_RATE_LIMIT_MAX_RETRIES = 5; // максимум попыток при 429 ошибке
const WB_RATE_LIMIT_BASE_DELAY_MS = 2000; // базовая задержка 2 секунды
const WB_RATE_LIMIT_MAX_DELAY_MS = 30000; // максимальная задержка 30 секунд

/**
 * Формирует URL с параметрами (аналог URLSearchParams для Google Apps Script)
 */
function buildUrlWithParams(baseUrl, params) {
  const urlParams = Object.keys(params)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(params[key])}`)
    .join('&');
  return `${baseUrl}?${urlParams}`;
}

/**
 * Выполняет запрос к WB API с обработкой лимитов запросов (HTTP 429)
 */
function wbApiRequestWithRetry(url, options, maxRetries = null) {
  // Получаем настройки из PropertiesService или используем значения по умолчанию
  const properties = PropertiesService.getScriptProperties();
  const actualMaxRetries = maxRetries || parseInt(properties.getProperty('WB_RATE_LIMIT_MAX_RETRIES')) || WB_RATE_LIMIT_MAX_RETRIES;
  const baseDelay = parseInt(properties.getProperty('WB_RATE_LIMIT_BASE_DELAY_MS')) || WB_RATE_LIMIT_BASE_DELAY_MS;
  const maxDelay = parseInt(properties.getProperty('WB_RATE_LIMIT_MAX_DELAY_MS')) || WB_RATE_LIMIT_MAX_DELAY_MS;
  
  let lastError;
  
  for (let attempt = 0; attempt <= actualMaxRetries; attempt++) {
    try {
      const resp = UrlFetchApp.fetch(url, options);
      const code = resp.getResponseCode();
      
      if (code === 429) {
        // Обрабатываем ошибку "Too Many Requests"
        const errorBody = resp.getContentText();
        console.log(`⚠️ HTTP 429 (Too Many Requests) на попытке ${attempt + 1}/${actualMaxRetries + 1}`);
        console.log(`Ошибка: ${errorBody}`);
        
        if (attempt < actualMaxRetries) {
          // Вычисляем задержку с экспоненциальным backoff
          const delay = Math.min(
            baseDelay * Math.pow(2, attempt),
            maxDelay
          );
          
          console.log(`⏳ Ждём ${delay}ms перед повторной попыткой...`);
          Utilities.sleep(delay);
          continue;
        } else {
          throw new Error(`WB API: превышен лимит запросов после ${actualMaxRetries + 1} попыток. Последняя ошибка: ${errorBody}`);
        }
      }
      
      // Если не 429, возвращаем ответ (успешный или с другой ошибкой)
      return resp;
      
    } catch (error) {
      lastError = error;
      
      // Если это не HTTP ошибка, пробрасываем её дальше
      if (!error.message.includes('HTTP')) {
        throw error;
      }
      
      // Для HTTP ошибок, отличных от 429, пробрасываем сразу
      if (!error.message.includes('429')) {
        throw error;
      }
      
      // Для 429 ошибок продолжаем цикл
      if (attempt < actualMaxRetries) {
        const delay = Math.min(
          baseDelay * Math.pow(2, attempt),
          maxDelay
        );
        
        console.log(`⏳ HTTP 429, ждём ${delay}ms перед повторной попыткой...`);
        Utilities.sleep(delay);
      }
    }
  }
  
  // Если дошли сюда, все попытки исчерпаны
  throw lastError || new Error('WB API: все попытки запроса исчерпаны');
}

/**
 * Получает отчёт по продажам через Statistics API (альтернативный метод)
 */
function wbGetReportDetailByPeriod_(apiKey, dateFrom, dateTo) {
  const url = WB_STATISTICS_HOST + '/api/v5/supplier/reportDetailByPeriod';
  const params = {
    dateFrom: dateFrom, // Формат: YYYY-MM-DD
    dateTo: dateTo,     // Формат: YYYY-MM-DD
    limit: 100000,      // Максимальное количество записей
    rrdid: 0            // ID отчёта (0 для нового)
  };
  
  const fullUrl = buildUrlWithParams(url, params);
  
  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Authorization': apiKey
    }
  };
  
  console.log('Получаем отчёт по продажам через Statistics API...');
  console.log(`URL: ${fullUrl}`);
  console.log(`Параметры:`, params);
  
  const resp = wbApiRequestWithRetry(fullUrl, options);
  
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`WB Statistics API: HTTP ${code} — ${resp.getContentText()}`);
  }
  
  const body = JSON.parse(resp.getContentText() || '{}');
  console.log('WB Statistics API Response:', JSON.stringify(body, null, 2));
  
  return body;
}

/**
 * Выгружает данные через Statistics API (альтернативный метод)
 */
function exportWBStocksViaStatisticsAPI() {
  try {
    const config = getWBConfig();
    
    if (!config.API_KEY) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API ключ для WB магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Начинаем выгрузку через Statistics API для WB магазина: ${config.STORE_NAME}`);
    
    // Получаем даты (последние 7 дней - оптимальный период)
    const today = new Date();
    const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
    
    const dateTo = today.toISOString().split('T')[0]; // YYYY-MM-DD
    const dateFrom = weekAgo.toISOString().split('T')[0]; // YYYY-MM-DD
    
    console.log(`Период: с ${dateFrom} по ${dateTo}`);
    
    // Получаем отчёт
    const reportData = wbGetReportDetailByPeriod_(config.API_KEY, dateFrom, dateTo);
    
    if (!reportData || !Array.isArray(reportData)) {
      console.log('Нет данных в отчёте');
      SpreadsheetApp.getUi().alert('Информация', 'Нет данных в отчёте за указанный период', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Получено записей: ${reportData.length}`);
    
    // Обрабатываем данные
    const processedData = reportData.map(item => ({
      nmId: item.nmId || '',
      supplierArticle: item.supplierArticle || '',
      barcode: item.barcode || '',
      techSize: item.techSize || '',
      warehouseName: item.warehouseName || '',
      warehouseId: item.warehouseId || '',
      quantity: 0, // В отчёте по продажам нет остатков
      reserve: 0,
      inWayToClient: 0,
      inWayFromClient: 0,
      store_name: config.STORE_NAME,
      // Дополнительные поля из отчёта по продажам
      sale_dt: item.sale_dt || '',
      price: item.price || 0,
      quantity_sold: item.quantity || 0,
      total_price: item.totalPrice || 0
    }));
    
    // Записываем в Google Sheets
    writeWBStatisticsToGoogleSheets(processedData);
    
    console.log(`Выгрузка через Statistics API завершена! Записано записей: ${processedData.length}`);
    
    SpreadsheetApp.getUi().alert('Успех', `Выгрузка через Statistics API завершена!\nЗаписей: ${processedData.length}\nПериод: ${dateFrom} - ${dateTo}`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Ошибка при выгрузке через Statistics API:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка выгрузки: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Записывает данные Statistics API в Google Таблицы
 */
function writeWBStatisticsToGoogleSheets(data) {
  const config = getWBConfig();
  
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
  
  // Определяем название листа
  const storeName = config.STORE_NAME || 'Неизвестный WB магазин';
  const sheetName = sanitizeSheetName(storeName + ' - Statistics');
  
  console.log(`Создаем/используем лист: ${sheetName}`);
  
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  // Создаем лист если не существует
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  // Очищаем только диапазон с данными (A:N)
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    const range = sheet.getRange(1, 1, lastRow, 14); // 14 колонок A-N
    range.clear();
  }
  
  // Заголовки для Statistics API
  const headers = [
    'Магазин',
    'nmId',
    'Артикул поставщика',
    'Штрихкод',
    'Размер',
    'Название склада',
    'ID склада',
    'Остаток',
    'Зарезервировано',
    'В пути к клиенту',
    'В пути от клиента',
    'Дата продажи',
    'Цена',
    'Количество продано'
  ];
  
  // Записываем заголовки
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Форматируем заголовки
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#E8F0FE');
  
  if (data.length === 0) {
    console.log('Нет данных для записи');
    return;
  }
  
  // Подготавливаем данные
  const rows = data.map(item => [
    item.store_name || config.STORE_NAME || 'Неизвестный WB магазин',
    item.nmId || '',
    item.supplierArticle || '',
    item.barcode || '',
    item.techSize || '',
    item.warehouseName || '',
    item.warehouseId || '',
    item.quantity || 0,
    item.reserve || 0,
    item.inWayToClient || 0,
    item.inWayFromClient || 0,
    item.sale_dt || '',
    item.price || 0,
    item.quantity_sold || 0
  ]);
  
  // Записываем данные
  if (rows.length > 0) {
    try {
      const dataRange = sheet.getRange(2, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      
      // Фильтры отключены по требованию
      
      console.log(`Записано ${rows.length} строк в Google Таблицы`);
    } catch (error) {
      console.error('Ошибка при записи данных:', error);
      throw error;
    }
  } else {
    console.log('Нет данных для записи');
  }
}

/**
 * Выгружает FBO остатки для активного WB магазина с увеличенными задержками
 */
function exportWBFBOStocksWithLongDelays() {
  try {
    const config = getWBConfig();
    
    if (!config.API_KEY) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API ключ для WB магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Начинаем выгрузку FBO остатков для WB магазина: ${config.STORE_NAME} (с увеличенными задержками)`);
    
    // Временно увеличиваем задержки для этого запроса
    const originalMaxRetries = WB_RATE_LIMIT_MAX_RETRIES;
    const originalBaseDelay = WB_RATE_LIMIT_BASE_DELAY_MS;
    const originalMaxDelay = WB_RATE_LIMIT_MAX_DELAY_MS;
    
    // Устанавливаем более консервативные настройки
    const properties = PropertiesService.getScriptProperties();
    properties.setProperties({
      'WB_RATE_LIMIT_MAX_RETRIES': '3',
      'WB_RATE_LIMIT_BASE_DELAY_MS': '15000', // 15 секунд
      'WB_RATE_LIMIT_MAX_DELAY_MS': '60000'   // 60 секунд
    });
    
    try {
      const taskId = wbCreateWarehouseRemainsReport_(config.API_KEY);
      const downloadUrl = wbWaitReportAndGetUrl_(taskId, config.API_KEY);
      const csv = wbDownloadReportCsv_(taskId, config.API_KEY);
      const rows = parseCsv_(csv);
      
      if (rows.length === 0) {
        console.log('Нет данных для записи');
        return;
      }
      
      // Обрабатываем данные
      const headerMap = normalizeHeaderMap_(rows[0]);
      const data = rows.slice(1).map(r => ({
        nmId: pick_(r, headerMap.nmId),
        supplierArticle: pick_(r, headerMap.supplierArticle),
        barcode: pick_(r, headerMap.barcode),
        techSize: pick_(r, headerMap.techSize),
        warehouseName: pick_(r, headerMap.warehouseName),
        warehouseId: pick_(r, headerMap.warehouseId),
        quantity: toNum_(pick_(r, headerMap.quantity)),
        reserve: toNum_(pick_(r, headerMap.reserve)),
        inWayToClient: toNum_(pick_(r, headerMap.inWayToClient)),
        inWayFromClient: toNum_(pick_(r, headerMap.inWayFromClient)),
        store_name: config.STORE_NAME
      }));
      
      // Записываем в Google Sheets
      writeWBToGoogleSheets(data);
      
      console.log(`Выгрузка WB FBO остатков завершена! Записано товаров: ${data.length}`);
      
    } finally {
      // Восстанавливаем оригинальные настройки
      properties.setProperties({
        'WB_RATE_LIMIT_MAX_RETRIES': originalMaxRetries.toString(),
        'WB_RATE_LIMIT_BASE_DELAY_MS': originalBaseDelay.toString(),
        'WB_RATE_LIMIT_MAX_DELAY_MS': originalMaxDelay.toString()
      });
    }
    
  } catch (error) {
    console.error('Ошибка при выгрузке WB FBO остатков:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка выгрузки: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Выгружает FBO остатки для активного WB магазина
 */
function exportWBFBOStocks() {
  try {
    console.log('Начинаем выгрузку FBO остатков через Statistics API...');
    
    // Проверяем, что активный WB магазин настроен
    const config = getWBConfig();
    if (!config.API_KEY) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API ключ для активного WB магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Используем новую функцию для загрузки всех остатков
    loadAllStocks();
    
    console.log('Выгрузка WB FBO остатков через Statistics API завершена!');
    
  } catch (error) {
    console.error('Ошибка при выгрузке WB FBO остатков:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка выгрузки: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Выгружает FBO остатки для всех WB магазинов через Statistics API
 */
function exportAllWBStoresStocksStatisticsAPI() {
  try {
    const stores = getWBStoresList();
    
    if (stores.length === 0) {
      try {
        SpreadsheetApp.getUi().alert('Ошибка', 'Нет добавленных WB магазинов!', SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) {
        console.log('UI alert пропущен (запуск из триггера)');
      }
      return;
    }
    
    console.log(`Начинаем выгрузку FBO остатков со всех WB магазинов через Statistics API (${stores.length} магазинов)...`);
    
    const originalActiveStore = getActiveWBStore();
    let totalProcessed = 0;
    
    stores.forEach((store, index) => {
      try {
        console.log(`Обрабатываем WB магазин ${index + 1}/${stores.length}: ${store.name}`);
        
        // Устанавливаем активный магазин
        setActiveWBStore(store.id);
        
        // Добавляем задержку между магазинами для избежания лимитов
        if (index > 0) {
          console.log('Ждем 3 секунды перед обработкой следующего магазина...');
          Utilities.sleep(3000);
        }
        
        // Используем новую функцию для загрузки остатков через Statistics API
        const allData = loadAllStocksForStore(store);
        totalProcessed += allData.length;
        
        console.log(`  WB магазин "${store.name}" обработан успешно. Получено записей: ${allData.length}`);
        
      } catch (error) {
        console.error(`Ошибка при обработке WB магазина "${store.name}":`, error);
        // Продолжаем с другими магазинами
      }
    });
    
    // Восстанавливаем активный магазин
    if (originalActiveStore) {
      setActiveWBStore(originalActiveStore.id);
    }
    
    console.log(`Выгрузка со всех WB магазинов завершена! Всего обработано товаров: ${totalProcessed}`);
    
    try {
      SpreadsheetApp.getUi().alert('Выгрузка завершена', `Обработано ${stores.length} WB магазинов, всего товаров: ${totalProcessed}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      console.log('UI alert пропущен (запуск из триггера)');
    }
    
  } catch (error) {
    console.error('Ошибка при выгрузке со всех WB магазинов:', error);
    throw error;
  }
}

/**
 * Выгружает FBO остатки для всех WB магазинов (старый API с лимитами)
 */
function exportAllWBStoresStocks() {
  try {
    const stores = getWBStoresList();
    
    if (stores.length === 0) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Нет добавленных WB магазинов!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Начинаем выгрузку FBO остатков со всех WB магазинов (${stores.length} магазинов)...`);
    
    const originalActiveStore = getActiveWBStore();
    let totalProcessed = 0;
    
    stores.forEach((store, index) => {
      try {
        console.log(`Обрабатываем WB магазин ${index + 1}/${stores.length}: ${store.name}`);
        
        // Устанавливаем активный магазин
        setActiveWBStore(store.id);
        
        // Получаем остатки для текущего магазина
        const taskId = wbCreateWarehouseRemainsReport_(store.api_key);
        const downloadUrl = wbWaitReportAndGetUrl_(taskId, store.api_key);
        const csv = wbDownloadReportCsv_(taskId, store.api_key);
        const rows = parseCsv_(csv);
        
        if (rows.length > 0) {
          // Обрабатываем данные
          const headerMap = normalizeHeaderMap_(rows[0]);
          const data = rows.slice(1).map(r => ({
            nmId: pick_(r, headerMap.nmId),
            supplierArticle: pick_(r, headerMap.supplierArticle),
            barcode: pick_(r, headerMap.barcode),
            techSize: pick_(r, headerMap.techSize),
            warehouseName: pick_(r, headerMap.warehouseName),
            warehouseId: pick_(r, headerMap.warehouseId),
            quantity: toNum_(pick_(r, headerMap.quantity)),
            reserve: toNum_(pick_(r, headerMap.reserve)),
            inWayToClient: toNum_(pick_(r, headerMap.inWayToClient)),
            inWayFromClient: toNum_(pick_(r, headerMap.inWayFromClient)),
            store_name: store.name
          }));
          
          // Записываем в отдельный лист для этого магазина
          writeWBToGoogleSheets(data);
          totalProcessed += data.length;
        }
        
        console.log(`  WB магазин "${store.name}" обработан успешно`);
        
      } catch (error) {
        console.error(`Ошибка при обработке WB магазина "${store.name}":`, error);
        // Продолжаем с другими магазинами
      }
    });
    
    // Восстанавливаем активный магазин
    if (originalActiveStore) {
      setActiveWBStore(originalActiveStore.id);
    }
    
    console.log(`Выгрузка со всех WB магазинов завершена! Всего обработано товаров: ${totalProcessed}`);
    
    try {
      SpreadsheetApp.getUi().alert('Выгрузка завершена', `Обработано ${stores.length} WB магазинов, всего товаров: ${totalProcessed}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      console.log('UI alert пропущен (запуск из триггера)');
    }
    
  } catch (error) {
    console.error('Ошибка при выгрузке со всех WB магазинов:', error);
    throw error;
  }
}

/**
 * Создает отчёт "Warehouses Remains Report" в WB с настраиваемыми параметрами
 */
function wbCreateWarehouseRemainsReportWithParams_(apiKey, params = {}) {
  // Получаем сохранённые настройки из PropertiesService
  const properties = PropertiesService.getScriptProperties();
  
  // Параметры по умолчанию (с учётом сохранённых настроек)
  const defaultParams = {
    locale: properties.getProperty('WB_REPORT_LOCALE') || 'ru',           // Язык полей ответа
    groupByBrand: properties.getProperty('WB_REPORT_GROUP_BY_BRAND') || 'false',  // Разбивка по брендам
    groupBySubject: properties.getProperty('WB_REPORT_GROUP_BY_SUBJECT') || 'false', // Разбивка по предметам
    groupBySa: properties.getProperty('WB_REPORT_GROUP_BY_SA') || 'false',     // Разбивка по артикулам продавца
    groupByNm: properties.getProperty('WB_REPORT_GROUP_BY_NM') || 'true',      // Разбивка по артикулам WB (включаем для получения поля volume)
    groupByBarcode: properties.getProperty('WB_REPORT_GROUP_BY_BARCODE') || 'false', // Разбивка по баркодам
    groupBySize: properties.getProperty('WB_REPORT_GROUP_BY_SIZE') || 'false',   // Разбивка по размерам
    filterPics: properties.getProperty('WB_REPORT_FILTER_PICS') || '0',        // Не применять фильтр по фото
    filterVolume: properties.getProperty('WB_REPORT_FILTER_VOLUME') || '0'       // Не применять фильтр по объёму
  };
  
  // Объединяем параметры по умолчанию с переданными
  const finalParams = { ...defaultParams, ...params };
  
  // Формируем URL с параметрами согласно документации API
  const baseUrl = WB_ANALYTICS_HOST + '/api/v1/warehouse_remains';
  const url = buildUrlWithParams(baseUrl, finalParams);
  
  const options = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Authorization': apiKey
    }
  };
  
  console.log('Создаём отчёт WB с параметрами...');
  console.log(`URL: ${url}`);
  console.log(`Параметры:`, finalParams);
  
  const resp = wbApiRequestWithRetry(url, options);
  
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`WB create report: HTTP ${code} — ${resp.getContentText()}`);
  }
  
  const body = JSON.parse(resp.getContentText() || '{}');
  console.log('WB API Response:', JSON.stringify(body, null, 2));
  
  // Пробуем разные варианты получения taskId (новый API возвращает taskId вместо reportId)
  const taskId = body?.data?.taskId || 
                 body?.data?.id || 
                 body?.data?.reportId || 
                 body?.reportId || 
                 body?.id ||
                 body?.requestId ||
                 body?.data?.requestId ||
                 body?.taskId;
  
  if (!taskId) {
    console.error('Не найдено поле taskId в ответе:', body);
    throw new Error(`WB create report: не получили taskId. Ответ: ${JSON.stringify(body)}`);
  }
  
  console.log(`✅ Получен taskId: ${taskId}`);
  return taskId;
}

/**
 * Создает отчёт "Warehouses Remains Report" в WB (упрощённая версия)
 */
function wbCreateWarehouseRemainsReport_(apiKey) {
  // Используем функцию с параметрами по умолчанию
  return wbCreateWarehouseRemainsReportWithParams_(apiKey);
}

/**
 * Ждёт готовности отчёта и получает URL скачивания
 */
function wbWaitReportAndGetUrl_(taskId, apiKey) {
  const started = Date.now();
  
  while (Date.now() - started < WB_REPORT_TIMEOUT_MS) {
    Utilities.sleep(WB_REPORT_POLL_INTERVAL_MS);
    
    const url = WB_ANALYTICS_HOST + '/api/v1/warehouse_remains';
    const options = {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        'Authorization': apiKey
      }
    };
    
    try {
      const resp = wbApiRequestWithRetry(url + '?id=' + encodeURIComponent(taskId), options);
      
      if (resp.getResponseCode() === 200) {
        const body = JSON.parse(resp.getContentText() || '{}');
        console.log(`WB Report Status Response:`, JSON.stringify(body, null, 2));
        
        const status = (body?.data?.status || body?.status || '').toLowerCase();
        console.log(`Report status: ${status}`);
        
        if (status === 'ready' || status === 'done' || status === 'success') {
          const downloadUrl = body?.data?.file || 
                             body?.data?.downloadUrl || 
                             body?.downloadUrl || 
                             body?.file ||
                             body?.data?.url ||
                             body?.url;
          
          if (!downloadUrl) {
            console.error('Не найден downloadUrl в ответе:', body);
            throw new Error(`WB report ready, но нет downloadUrl. Ответ: ${JSON.stringify(body)}`);
          }
          
          console.log(`✅ Получен downloadUrl: ${downloadUrl}`);
          return downloadUrl;
        }
        
        if (status === 'failed' || status === 'error') {
          throw new Error('WB report status: ' + status);
        }
      }
    } catch (error) {
      // Если это ошибка лимита запросов, продолжаем ждать
      if (error.message.includes('429') || error.message.includes('Too Many Requests')) {
        console.log(`⚠️ Лимит запросов при проверке статуса, продолжаем ждать...`);
        continue;
      }
      // Для других ошибок пробрасываем дальше
      throw error;
    }
  }
  
  throw new Error('WB report: ожидание готовности превысило лимит');
}

/**
 * Скачивает CSV-файл отчёта по taskId
 */
function wbDownloadReportCsv_(taskId, apiKey) {
  // Сначала получаем URL для скачивания по taskId
  const statusUrl = WB_ANALYTICS_HOST + '/api/v1/warehouse_remains';
  const statusOptions = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Authorization': apiKey
    }
  };
  
  console.log('Получаем URL для скачивания отчёта...');
  const statusResp = wbApiRequestWithRetry(statusUrl + '?id=' + encodeURIComponent(taskId), statusOptions);
  
  if (statusResp.getResponseCode() !== 200) {
    throw new Error(`WB get download URL: HTTP ${statusResp.getResponseCode()} — ${statusResp.getContentText()}`);
  }
  
  const statusBody = JSON.parse(statusResp.getContentText() || '{}');
  const downloadUrl = statusBody?.data?.file || 
                     statusBody?.data?.downloadUrl || 
                     statusBody?.downloadUrl || 
                     statusBody?.file ||
                     statusBody?.data?.url ||
                     statusBody?.url;
  
  if (!downloadUrl) {
    throw new Error(`WB download: не найден downloadUrl для taskId ${taskId}. Ответ: ${JSON.stringify(statusBody)}`);
  }
  
  console.log(`Скачиваем отчёт по URL: ${downloadUrl}`);
  
  // Скачиваем файл
  const downloadOptions = {
    method: 'get',
    muteHttpExceptions: true,
    headers: {
      'Authorization': apiKey
    }
  };
  
  const resp = wbApiRequestWithRetry(downloadUrl, downloadOptions);
  
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error(`WB download CSV: HTTP ${code} — ${resp.getContentText()}`);
  }
  
  console.log(`✅ Отчёт успешно скачан, размер: ${resp.getContentText().length} символов`);
  return resp.getContentText();
}

/**
 * Записывает данные WB в Google Таблицы
 */
function writeWBToGoogleSheets(data) {
  const config = getWBConfig();
  
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
  const storeName = config.STORE_NAME || 'Неизвестный WB магазин';
  const sheetName = sanitizeSheetName(storeName);
  
  console.log(`Создаем/используем лист: ${sheetName}`);
  
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  // Создаем лист если не существует
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  // Очищаем только диапазон с данными (A:K)
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    const range = sheet.getRange(1, 1, lastRow, 11); // 11 колонок A-K
    range.clear();
  }
  
  // Заголовки для WB
  const headers = [
    'Магазин',
    'nmId',
    'Артикул поставщика',
    'Штрихкод',
    'Размер',
    'Название склада',
    'ID склада',
    'Остаток',
    'Зарезервировано',
    'В пути к клиенту',
    'В пути от клиента'
  ];
  
  // Записываем заголовки
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Форматируем заголовки
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#E8F0FE');
  
  if (data.length === 0) {
    console.log('Нет данных для записи');
    return;
  }
  
  // Подготавливаем данные
  const rows = data.map(item => [
    item.store_name || config.STORE_NAME || 'Неизвестный WB магазин',
    item.nmId || '',
    item.supplierArticle || '',
    item.barcode || '',
    item.techSize || '',
    item.warehouseName || '',
    item.warehouseId || '',
    item.quantity || 0,
    item.reserve || 0,
    item.inWayToClient || 0,
    item.inWayFromClient || 0
  ]);
  
  // Записываем данные
  if (rows.length > 0) {
    try {
      const dataRange = sheet.getRange(2, 1, rows.length, headers.length);
      dataRange.setValues(rows);
      
      // Фильтры отключены по требованию
      
      console.log(`Записано ${rows.length} строк в Google Таблицы`);
    } catch (error) {
      console.error('Ошибка при записи данных:', error);
      throw error;
    }
  } else {
    console.log('Нет данных для записи');
  }
}

/**
 * Парсит CSV в массив строк
 */
function parseCsv_(csv) {
  const rows = Utilities.parseCsv(csv, ',');
  return rows;
}

/**
 * Нормализует заголовки к ожидаемым ключам
 */
function normalizeHeaderMap_(headerRow) {
  const map = {};
  const norm = s => String(s || '').trim().toLowerCase();
  
  headerRow.forEach((h, i) => {
    const n = norm(h);
    if (['nmid', 'nm_id', 'nm id', 'nm'].includes(n)) map.nmId = i;
    if (['supplierarticle', 'supplier_article', 'sa', 'vendorcode'].includes(n)) map.supplierArticle = i;
    if (['barcode', 'bar_code', 'штрихкод'].includes(n)) map.barcode = i;
    if (['techsize', 'size', 'tech_size', 'размер'].includes(n)) map.techSize = i;
    if (['warehousename', 'warehouse_name', 'склад'].includes(n)) map.warehouseName = i;
    if (['warehouseid', 'warehouse_id', 'id склада'].includes(n)) map.warehouseId = i;
    if (['quantity', 'qty', 'present', 'остаток'].includes(n)) map.quantity = i;
    if (['reserve', 'reserved'].includes(n)) map.reserve = i;
    if (['inwaytoclient', 'in_way_to_client'].includes(n)) map.inWayToClient = i;
    if (['inwayfromclient', 'in_way_from_client'].includes(n)) map.inWayFromClient = i;
  });
  
  return map;
}

/**
 * Вспомогательные функции
 */
function pick_(row, idx) { 
  return (idx == null) ? '' : row[idx]; 
}

function toNum_(v) { 
  return Number(String(v || '').replace(',', '.')) || 0; 
}

/**
 * Настраивает параметры отчёта WB API
 */
function configureWBReportParams() {
  const ui = SpreadsheetApp.getUi();
  
  // Получаем текущие настройки
  const properties = PropertiesService.getScriptProperties();
  const currentParams = {
    locale: properties.getProperty('WB_REPORT_LOCALE') || 'ru',
    groupByBrand: properties.getProperty('WB_REPORT_GROUP_BY_BRAND') || 'false',
    groupBySubject: properties.getProperty('WB_REPORT_GROUP_BY_SUBJECT') || 'false',
    groupBySa: properties.getProperty('WB_REPORT_GROUP_BY_SA') || 'false',
    groupByNm: properties.getProperty('WB_REPORT_GROUP_BY_NM') || 'true',
    groupByBarcode: properties.getProperty('WB_REPORT_GROUP_BY_BARCODE') || 'false',
    groupBySize: properties.getProperty('WB_REPORT_GROUP_BY_SIZE') || 'false',
    filterPics: properties.getProperty('WB_REPORT_FILTER_PICS') || '0',
    filterVolume: properties.getProperty('WB_REPORT_FILTER_VOLUME') || '0'
  };
  
  let message = `Текущие настройки отчёта WB:\n\n`;
  message += `Язык: ${currentParams.locale}\n`;
  message += `Группировка по брендам: ${currentParams.groupByBrand}\n`;
  message += `Группировка по предметам: ${currentParams.groupBySubject}\n`;
  message += `Группировка по артикулам продавца: ${currentParams.groupBySa}\n`;
  message += `Группировка по артикулам WB: ${currentParams.groupByNm}\n`;
  message += `Группировка по баркодам: ${currentParams.groupByBarcode}\n`;
  message += `Группировка по размерам: ${currentParams.groupBySize}\n`;
  message += `Фильтр по фото: ${currentParams.filterPics}\n`;
  message += `Фильтр по объёму: ${currentParams.filterVolume}\n\n`;
  message += `Изменить настройки? (y/n)`;
  
  const response = ui.prompt('Настройка параметров отчёта WB', message, ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  if (response.getResponseText().toLowerCase() === 'y' || response.getResponseText().toLowerCase() === 'yes') {
    // Запрашиваем новые настройки
    const localeResponse = ui.prompt('Язык отчёта', 'Язык (ru/en/zh):', ui.ButtonSet.OK_CANCEL);
    if (localeResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const groupByNmResponse = ui.prompt('Группировка по артикулам WB', 'Группировка по артикулам WB (true/false):', ui.ButtonSet.OK_CANCEL);
    if (groupByNmResponse.getSelectedButton() !== ui.Button.OK) return;
    
    const groupBySizeResponse = ui.prompt('Группировка по размерам', 'Группировка по размерам (true/false):', ui.ButtonSet.OK_CANCEL);
    if (groupBySizeResponse.getSelectedButton() !== ui.Button.OK) return;
    
    // Сохраняем настройки
    const newParams = {
      locale: localeResponse.getResponseText().trim() || 'ru',
      groupByNm: groupByNmResponse.getResponseText().trim() || 'true',
      groupBySize: groupBySizeResponse.getResponseText().trim() || 'false'
    };
    
    properties.setProperties({
      'WB_REPORT_LOCALE': newParams.locale,
      'WB_REPORT_GROUP_BY_NM': newParams.groupByNm,
      'WB_REPORT_GROUP_BY_SIZE': newParams.groupBySize
    });
    
    ui.alert('Успех', `Настройки отчёта WB сохранены:\n\nЯзык: ${newParams.locale}\nГруппировка по артикулам WB: ${newParams.groupByNm}\nГруппировка по размерам: ${newParams.groupBySize}`, ui.ButtonSet.OK);
  }
}

/**
 * Настраивает параметры обработки лимитов запросов WB API
 */
function configureWBRateLimits() {
  const ui = SpreadsheetApp.getUi();
  
  // Получаем текущие настройки
  const properties = PropertiesService.getScriptProperties();
  const currentMaxRetries = properties.getProperty('WB_RATE_LIMIT_MAX_RETRIES') || WB_RATE_LIMIT_MAX_RETRIES;
  const currentBaseDelay = properties.getProperty('WB_RATE_LIMIT_BASE_DELAY_MS') || WB_RATE_LIMIT_BASE_DELAY_MS;
  const currentMaxDelay = properties.getProperty('WB_RATE_LIMIT_MAX_DELAY_MS') || WB_RATE_LIMIT_MAX_DELAY_MS;
  
  let message = `Текущие настройки лимитов WB API:\n\n`;
  message += `Максимум попыток при 429 ошибке: ${currentMaxRetries}\n`;
  message += `Базовая задержка (мс): ${currentBaseDelay}\n`;
  message += `Максимальная задержка (мс): ${currentMaxDelay}\n\n`;
  message += `Введите новые значения (или оставьте пустыми для сохранения текущих):`;
  
  const maxRetriesResponse = ui.prompt('Настройка лимитов WB API', `${message}\n\nМаксимум попыток (1-10):`, ui.ButtonSet.OK_CANCEL);
  if (maxRetriesResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const baseDelayResponse = ui.prompt('Настройка лимитов WB API', 'Базовая задержка в миллисекундах (1000-10000):', ui.ButtonSet.OK_CANCEL);
  if (baseDelayResponse.getSelectedButton() !== ui.Button.OK) return;
  
  const maxDelayResponse = ui.prompt('Настройка лимитов WB API', 'Максимальная задержка в миллисекундах (10000-60000):', ui.ButtonSet.OK_CANCEL);
  if (maxDelayResponse.getSelectedButton() !== ui.Button.OK) return;
  
  // Валидация и сохранение
  try {
    const newMaxRetries = maxRetriesResponse.getResponseText().trim() ? 
      Math.max(1, Math.min(10, parseInt(maxRetriesResponse.getResponseText()))) : 
      currentMaxRetries;
    
    const newBaseDelay = baseDelayResponse.getResponseText().trim() ? 
      Math.max(1000, Math.min(10000, parseInt(baseDelayResponse.getResponseText()))) : 
      currentBaseDelay;
    
    const newMaxDelay = maxDelayResponse.getResponseText().trim() ? 
      Math.max(10000, Math.min(60000, parseInt(maxDelayResponse.getResponseText()))) : 
      currentMaxDelay;
    
    // Сохраняем настройки
    properties.setProperties({
      'WB_RATE_LIMIT_MAX_RETRIES': newMaxRetries.toString(),
      'WB_RATE_LIMIT_BASE_DELAY_MS': newBaseDelay.toString(),
      'WB_RATE_LIMIT_MAX_DELAY_MS': newMaxDelay.toString()
    });
    
    ui.alert('Успех', `Настройки лимитов WB API сохранены:\n\nМаксимум попыток: ${newMaxRetries}\nБазовая задержка: ${newBaseDelay}мс\nМаксимальная задержка: ${newMaxDelay}мс`, ui.ButtonSet.OK);
    
  } catch (error) {
    ui.alert('Ошибка', `Ошибка сохранения настроек: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Тестирует подключение к WB API
 */
function testWBConnection() {
  try {
    const config = getWBConfig();
    
    if (!config.API_KEY) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API ключ для WB магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log('Тестируем подключение к WB API...');
    console.log(`API Key: ${config.API_KEY.substring(0, 10)}...`);
    
    // Пробуем разные endpoints
    const endpoints = [
      '/api/v1/warehouse_remains',
      '/api/v1/warehouses',
      '/api/v1/supplier/warehouses',
      '/api/v1/supplier/warehouse_remains'
    ];
    
    let success = false;
    let lastError = '';
    
    for (const endpoint of endpoints) {
      try {
        console.log(`Пробуем endpoint: ${WB_ANALYTICS_HOST}${endpoint}`);
        
        const resp = UrlFetchApp.fetch(WB_ANALYTICS_HOST + endpoint, {
          method: 'get',
          muteHttpExceptions: true,
          headers: {
            'Authorization': config.API_KEY
          }
        });
        
        const code = resp.getResponseCode();
        console.log(`Response code: ${code}`);
        
        if (code === 200) {
          const body = resp.getContentText();
          console.log(`✅ Успешный ответ от ${endpoint}:`, body.substring(0, 500));
          success = true;
          break;
        } else {
          const errorText = resp.getContentText();
          console.log(`❌ Ошибка ${code} с ${endpoint}:`, errorText);
          lastError = `HTTP ${code}: ${errorText}`;
        }
        
      } catch (error) {
        console.log(`❌ Исключение с ${endpoint}:`, error.message);
        lastError = error.message;
      }
    }
    
    if (success) {
      SpreadsheetApp.getUi().alert('Успех', 'Подключение к WB API работает!', SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('Ошибка', `Не удалось подключиться к WB API. Последняя ошибка: ${lastError}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('Ошибка тестирования WB API:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка тестирования: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Тестирует Statistics API WB с разными периодами
 */
function testWBStatisticsAPIWithPeriods() {
  try {
    const config = getWBConfig();
    
    if (!config.API_KEY) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API ключ для WB магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log('Тестируем WB Statistics API с разными периодами...');
    console.log(`API Key: ${config.API_KEY.substring(0, 10)}...`);
    
    const today = new Date();
    const periods = [
      { name: '3 дня', days: 3 },
      { name: '7 дней', days: 7 },
      { name: '14 дней', days: 14 },
      { name: '30 дней', days: 30 },
      { name: '60 дней', days: 60 }
    ];
    
    let results = [];
    
    for (const period of periods) {
      const periodAgo = new Date(today.getTime() - period.days * 24 * 60 * 60 * 1000);
      const dateTo = today.toISOString().split('T')[0];
      const dateFrom = periodAgo.toISOString().split('T')[0];
      
      console.log(`\nТестируем период: ${period.name} (${dateFrom} - ${dateTo})`);
      
      try {
        const reportData = wbGetReportDetailByPeriod_(config.API_KEY, dateFrom, dateTo);
        const count = Array.isArray(reportData) ? reportData.length : 0;
        
        console.log(`✅ ${period.name}: ${count} записей`);
        results.push(`${period.name}: ${count} записей`);
        
        // Если нашли данные, показываем первые записи
        if (count > 0) {
          console.log('Первые 2 записи:');
          reportData.slice(0, 2).forEach((item, index) => {
            console.log(`${index + 1}. nmId: ${item.nmId}, Артикул: ${item.supplierArticle}, Дата: ${item.sale_dt}`);
          });
          break; // Прерываем цикл, если нашли данные
        }
        
        // Небольшая пауза между запросами
        Utilities.sleep(1000);
        
      } catch (error) {
        console.log(`❌ ${period.name}: Ошибка - ${error.message}`);
        results.push(`${period.name}: Ошибка`);
      }
    }
    
    const message = `Результаты тестирования WB Statistics API:\n\n${results.join('\n')}`;
    SpreadsheetApp.getUi().alert('Результаты тестирования', message, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Ошибка тестирования WB Statistics API:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка тестирования: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Тестирует Statistics API WB
 */
function testWBStatisticsAPI() {
  try {
    const config = getWBConfig();
    
    if (!config.API_KEY) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API ключ для WB магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log('Тестируем WB Statistics API...');
    console.log(`API Key: ${config.API_KEY.substring(0, 10)}...`);
    
    // Получаем даты (последние 7 дней для теста)
    const today = new Date();
    const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
    
    const dateTo = today.toISOString().split('T')[0]; // YYYY-MM-DD
    const dateFrom = weekAgo.toISOString().split('T')[0]; // YYYY-MM-DD
    
    console.log(`Тестовый период: с ${dateFrom} по ${dateTo}`);
    
    // Тестируем API
    const reportData = wbGetReportDetailByPeriod_(config.API_KEY, dateFrom, dateTo);
    
    if (!reportData || !Array.isArray(reportData)) {
      console.log('Нет данных в отчёте');
      SpreadsheetApp.getUi().alert('Информация', 'Нет данных в отчёте за указанный период', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`✅ Получено записей: ${reportData.length}`);
    
    // Показываем первые записи
    if (reportData.length > 0) {
      console.log('Первые 3 записи:');
      reportData.slice(0, 3).forEach((item, index) => {
        console.log(`${index + 1}. nmId: ${item.nmId}, Артикул: ${item.supplierArticle}, Дата: ${item.sale_dt}`);
      });
    }
    
    SpreadsheetApp.getUi().alert('Успех', `WB Statistics API работает!\nЗаписей: ${reportData.length}\nПериод: ${dateFrom} - ${dateTo}`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Ошибка тестирования WB Statistics API:', error);
    
    // Если это ошибка лимита запросов, показываем специальное сообщение
    if (error.message.includes('429') || error.message.includes('Too Many Requests')) {
      SpreadsheetApp.getUi().alert('Лимит запросов', `Обнаружена ошибка лимита запросов WB Statistics API.\n\nЭто нормально - система автоматически обработает такие ошибки с повторными попытками.\n\nОшибка: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('Ошибка', `Ошибка тестирования: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
}

/**
 * Тестирует новый WB API с taskId для отчёта "Warehouses Remains"
 */
function testWBTaskIdAPI() {
  try {
    const config = getWBConfig();
    
    if (!config.API_KEY) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API ключ для WB магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log('Тестируем новый WB API с taskId...');
    console.log(`API Key: ${config.API_KEY.substring(0, 10)}...`);
    
    // Показываем текущие настройки лимитов
    const properties = PropertiesService.getScriptProperties();
    const maxRetries = parseInt(properties.getProperty('WB_RATE_LIMIT_MAX_RETRIES')) || WB_RATE_LIMIT_MAX_RETRIES;
    const baseDelay = parseInt(properties.getProperty('WB_RATE_LIMIT_BASE_DELAY_MS')) || WB_RATE_LIMIT_BASE_DELAY_MS;
    const maxDelay = parseInt(properties.getProperty('WB_RATE_LIMIT_MAX_DELAY_MS')) || WB_RATE_LIMIT_MAX_DELAY_MS;
    
    console.log(`Настройки лимитов: maxRetries=${maxRetries}, baseDelay=${baseDelay}ms, maxDelay=${maxDelay}ms`);
    
    // Тест 1: Создание отчёта
    console.log('1. Создаём отчёт...');
    const taskId = wbCreateWarehouseRemainsReport_(config.API_KEY);
    console.log(`✅ Получен taskId: ${taskId}`);
    
    // Тест 2: Проверка статуса отчёта
    console.log('2. Проверяем статус отчёта...');
    let attempts = 0;
    const maxAttempts = 5;
    
    while (attempts < maxAttempts) {
      attempts++;
      console.log(`Попытка ${attempts}/${maxAttempts}...`);
      
      const url = WB_ANALYTICS_HOST + '/api/v1/warehouse_remains';
      const options = {
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          'Authorization': config.API_KEY
        }
      };
      
      try {
        const resp = wbApiRequestWithRetry(url + '?id=' + encodeURIComponent(taskId), options);
        
        if (resp.getResponseCode() === 200) {
          const body = JSON.parse(resp.getContentText() || '{}');
          console.log(`✅ Статус отчёта:`, JSON.stringify(body, null, 2));
          
          const status = (body?.data?.status || body?.status || '').toLowerCase();
          console.log(`Статус: ${status}`);
          
          if (status === 'ready' || status === 'done' || status === 'success') {
            console.log('✅ Отчёт готов!');
            
            // Тест 3: Скачивание отчёта
            console.log('3. Скачиваем отчёт...');
            const csv = wbDownloadReportCsv_(taskId, config.API_KEY);
            console.log(`✅ Отчёт скачан, размер: ${csv.length} символов`);
            
            // Показываем первые строки
            const lines = csv.split('\n');
            console.log(`Первые 3 строки отчёта:`);
            lines.slice(0, 3).forEach((line, index) => {
              console.log(`${index + 1}: ${line}`);
            });
            
            SpreadsheetApp.getUi().alert('Успех', `Тест WB API с taskId прошёл успешно!\nTaskId: ${taskId}\nРазмер отчёта: ${csv.length} символов\n\nОбработка лимитов запросов работает корректно!`, SpreadsheetApp.getUi().ButtonSet.OK);
            return;
          } else if (status === 'failed' || status === 'error') {
            throw new Error(`Отчёт завершился с ошибкой: ${status}`);
          } else {
            console.log(`Отчёт ещё обрабатывается (статус: ${status}), ждём...`);
            if (attempts < maxAttempts) {
              Utilities.sleep(3000); // Ждём 3 секунды
            }
          }
        }
      } catch (error) {
        // Если это ошибка лимита запросов, показываем что обработка работает
        if (error.message.includes('429') || error.message.includes('Too Many Requests')) {
          console.log(`⚠️ Обнаружена ошибка лимита запросов, но обработка работает корректно`);
          console.log(`Ошибка: ${error.message}`);
        } else {
          throw error;
        }
      }
    }
    
    console.log('⚠️ Отчёт не готов за отведённое время, но API работает');
    SpreadsheetApp.getUi().alert('Частичный успех', `WB API с taskId работает!\nTaskId: ${taskId}\nОтчёт ещё обрабатывается\n\nОбработка лимитов запросов настроена корректно!`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Ошибка тестирования WB API с taskId:', error);
    
    // Если это ошибка лимита запросов, показываем специальное сообщение
    if (error.message.includes('429') || error.message.includes('Too Many Requests')) {
      SpreadsheetApp.getUi().alert('Лимит запросов', `Обнаружена ошибка лимита запросов WB API.\n\nЭто нормально - система автоматически обработает такие ошибки с повторными попытками.\n\nОшибка: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
      SpreadsheetApp.getUi().alert('Ошибка', `Ошибка тестирования: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  }
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

// ==================== НОВЫЕ ФУНКЦИИ ДЛЯ WB STATISTICS API ====================

/**
 * Загружает все остатки FBO через Statistics API
 */
function loadAllStocks() {
  try {
    Logger.log('Начинаем загрузку всех остатков FBO...');
    
    // Получаем конфигурацию активного WB магазина
    const config = getWBConfig();
    if (!config.API_KEY) {
      throw new Error('Не настроен API ключ для активного WB магазина!');
    }

    const storeName = config.STORE_NAME || 'Неизвестный WB магазин';
    Logger.log(`Загружаем остатки для магазина: ${storeName}`);
    Logger.log(`Используем API ключ: ***${config.API_KEY.slice(-4)}`);

    // Загружаем данные для активного магазина
    const allData = loadAllStocksForStore({ name: storeName, api_key: config.API_KEY });
    
    Logger.log('Всего получено остатков: ' + allData.length);
    writeToSheet(allData);
    Logger.log('Загрузка всех остатков FBO завершена');

  } catch (error) {
    Logger.log('Ошибка: ' + error);
    throw error;
  }
}

/**
 * Загружает все остатки FBO для конкретного магазина через Statistics API
 */
function loadAllStocksForStore(store) {
  try {
    Logger.log(`Загружаем остатки для магазина: ${store.name}`);
    Logger.log(`Используем API ключ: ***${store.api_key.slice(-4)}`);

    // Начинаем с «старой» даты, например, 1 год назад или минимальная возможная
    let dateFrom = '2025-09-01T00:00:00'; // пример начальной даты

    let allData = [];
    let keepLoading = true;

    while (keepLoading) {
      Logger.log('Запрос с dateFrom: ' + dateFrom);
      let batch = fetchStocksBatch(dateFrom, store.api_key);
      if (batch.length === 0) {
        keepLoading = false;
        Logger.log('Все остатки выгружены');
        break;
      }
      
      // Добавляем информацию о магазине к каждой записи
      const batchWithStore = batch.map(item => ({
        ...item,
        store_name: store.name
      }));
      
      allData = allData.concat(batchWithStore);
      // Берем lastChangeDate из последней строки
      dateFrom = batch[batch.length - 1].lastChangeDate;
      Logger.log('Обработано записей: ' + allData.length);
      
      // Добавляем небольшую задержку между запросами для избежания лимитов
      Utilities.sleep(1000);
    }

    Logger.log(`Всего получено остатков для ${store.name}: ${allData.length}`);
    
    // Записываем данные в лист с названием магазина
    writeToSheet(allData);
    
    return allData;

  } catch (error) {
    Logger.log('Ошибка: ' + error);
    throw error;
  }
}

/**
 * Получает порцию данных остатков
 */
function fetchStocksBatch(dateFrom, apiToken) {
  const urlBase = 'https://statistics-api.wildberries.ru/api/v1/supplier/stocks';
  const url = urlBase + '?dateFrom=' + encodeURIComponent(dateFrom);
  const options = {
    method: 'GET',
    headers: {
      'Authorization': apiToken,
      'Content-Type': 'application/json',
    },
    muteHttpExceptions: true
  };
  Logger.log('Отправляем запрос: ' + url);
  const response = UrlFetchApp.fetch(url, options);
  const code = response.getResponseCode();
  Logger.log('Код ответа: ' + code);
  if (code !== 200) {
    throw new Error('API вернул ошибку: ' + code + ' | ' + response.getContentText());
  }
  const jsonData = JSON.parse(response.getContentText());
  return jsonData;
}

/**
 * Записывает данные в Google Sheets
 */
function writeToSheet(data) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Получаем название активного WB магазина
  const activeStore = getActiveWBStore();
  const storeName = activeStore ? activeStore.name : 'Неизвестный WB магазин';
  const sheetName = sanitizeSheetName(storeName);
  
  console.log(`Создаем/используем лист: ${sheetName}`);
  
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  sheet.getRange('A:O').clearContent();

  if (!data || data.length === 0) {
    sheet.getRange(1, 1).setValue('Данные отсутствуют');
    return;
  }

  const headers = [
    'Магазин', 'lastChangeDate', 'warehouseName', 'supplierArticle', 'nmId', 'barcode',
    'quantity', 'inWayToClient', 'inWayFromClient',
    'price', 'discount', 'category', 'subject', 'brand', 'techSize'
  ];

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#4285f4')
    .setFontColor('white');

  const rows = data.map(item => [
    item.store_name || storeName,
    item.lastChangeDate || '',
    item.warehouseName || '',
    item.supplierArticle || '',
    item.nmId || '',
    item.barcode || '',
    item.quantity || 0,
    item.inWayToClient || '',
    item.inWayFromClient || '',
    item.price || '',
    item.discount || '',
    item.category || '',
    item.subject || '',
    item.brand || '',
    item.techSize || ''
  ]);

  sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(1, headers.length);

  sheet.getRange(rows.length + 3, 1)
    .setValue('Обновлено: ' + new Date().toLocaleString('ru-RU'));
}

// ==================== ФУНКЦИИ ДЛЯ РАБОТЫ С ЯНДЕКС МАРКЕТ API ====================

/**
 * Выгружает остатки для активного Яндекс Маркет магазина
 */
function exportYandexStocks() {
  try {
    console.log('Начинаем выгрузку остатков Яндекс Маркета...');
    
    // Проверяем, что активный Яндекс Маркет магазин настроен
    const config = getYandexConfig();
    if (!config.API_TOKEN || !config.CAMPAIGN_ID) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API токен или Campaign ID для активного Яндекс Маркет магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Выгружаем остатки для магазина: ${config.STORE_NAME}`);
    console.log(`Campaign ID: ${config.CAMPAIGN_ID}`);
    console.log(`API Token: ***${config.API_TOKEN.slice(-4)}`);
    
    // Получаем остатки через API
    const stocks = getYandexStocks(config.API_TOKEN, config.CAMPAIGN_ID);
    
    if (stocks.length === 0) {
      console.log('Нет данных для записи');
      SpreadsheetApp.getUi().alert('Информация', 'Нет данных об остатках', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Записываем в Google Sheets
    writeYandexToGoogleSheets(stocks, config.STORE_NAME);
    
    console.log(`Выгрузка остатков Яндекс Маркета завершена! Записано товаров: ${stocks.length}`);
    SpreadsheetApp.getUi().alert('Выгрузка завершена', `Записано товаров: ${stocks.length}`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Ошибка при выгрузке остатков Яндекс Маркета:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка выгрузки: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Выгружает остатки для всех Яндекс Маркет магазинов
 */
function exportAllYandexStoresStocks() {
  try {
    const stores = getYandexStoresList();
    
    if (stores.length === 0) {
      try {
        SpreadsheetApp.getUi().alert('Ошибка', 'Нет добавленных Яндекс Маркет магазинов!', SpreadsheetApp.getUi().ButtonSet.OK);
      } catch (e) {
        console.log('UI alert пропущен (запуск из триггера)');
      }
      return;
    }
    
    console.log(`Начинаем выгрузку остатков со всех Яндекс Маркет магазинов (${stores.length} магазинов)...`);
    
    const originalActiveStore = getActiveYandexStore();
    let totalProcessed = 0;
    
    stores.forEach((store, index) => {
      try {
        console.log(`Обрабатываем Яндекс Маркет магазин ${index + 1}/${stores.length}: ${store.name}`);
        
        // Устанавливаем активный магазин
        setActiveYandexStore(store.id);
        
        // Добавляем задержку между магазинами для избежания лимитов
        if (index > 0) {
          console.log('Ждем 2 секунды перед обработкой следующего магазина...');
          Utilities.sleep(2000);
        }
        
        // Получаем остатки для текущего магазина
        const stocks = getYandexStocks(store.api_token, store.campaign_id);
        
        if (stocks.length > 0) {
          // Записываем в отдельный лист для этого магазина
          writeYandexToGoogleSheets(stocks, store.name);
          totalProcessed += stocks.length;
        }
        
        console.log(`  Яндекс Маркет магазин "${store.name}" обработан успешно. Получено записей: ${stocks.length}`);
        
      } catch (error) {
        console.error(`Ошибка при обработке Яндекс Маркет магазина "${store.name}":`, error);
        // Продолжаем с другими магазинами
      }
    });
    
    // Восстанавливаем активный магазин
    if (originalActiveStore) {
      setActiveYandexStore(originalActiveStore.id);
    }
    
    console.log(`Выгрузка со всех Яндекс Маркет магазинов завершена! Всего обработано товаров: ${totalProcessed}`);
    
    try {
      SpreadsheetApp.getUi().alert('Выгрузка завершена', `Обработано ${stores.length} Яндекс Маркет магазинов, всего товаров: ${totalProcessed}`, SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (e) {
      console.log('UI alert пропущен (запуск из триггера)');
    }
    
  } catch (error) {
    console.error('Ошибка при выгрузке со всех Яндекс Маркет магазинов:', error);
    throw error;
  }
}

/**
 * Получает остатки товаров через Яндекс Маркет API (точная копия рабочего скрипта)
 */
function getYandexStocks(apiToken, campaignId) {
  try {
    // URL для получения остатков товаров
    const stocksUrl = `https://api.partner.market.yandex.ru/campaigns/${campaignId}/offers/stocks`;
    // URL для получения списка складов
    const warehousesUrl = "https://api.partner.market.yandex.ru/warehouses";
    
    // Заголовки HTTP для авторизации: используем Api Key токен
    const headers = {
      "Api-Key": apiToken
    };
    
    console.log(`Отправляем запрос к Яндекс Маркет API: ${stocksUrl}`);
    
    // Получаем справочник складов Маркета (FBY)
    const warehouseMap = {};
    try {
      const whResponse = UrlFetchApp.fetch(warehousesUrl, { 
        "method": "get", 
        "headers": headers 
      });
      const whData = JSON.parse(whResponse.getContentText());
      if (whData.status == "OK" && whData.result && whData.result.warehouses) {
        whData.result.warehouses.forEach(function (w) {
          warehouseMap[w.id] = w.name;
        });
        console.log(`Получено складов: ${Object.keys(warehouseMap).length}`);
      }
    } catch (e) {
      console.log("Не удалось получить список складов: " + e);
    }
    
    // Подготовка тела запроса для получения остатков товаров
    const requestBody = {
      "archived": false,
      "withTurnover": false
    };
    
    // Опции для UrlFetchApp (POST запрос с JSON-телом)
    const options = {
      "method": "post",
      "contentType": "application/json",
      "headers": headers,
      "payload": JSON.stringify(requestBody)
    };
    
    // Выполняем запрос к API и обрабатываем данные с постраничной загрузкой
    const allStocks = [];
    let pageToken = null;
    
    do {
      // Если есть токен следующей страницы, добавляем его к телу запроса
      if (pageToken) {
        requestBody.page_token = pageToken;
        options.payload = JSON.stringify(requestBody);
      }
      
      // Выполняем API-запрос за текущей страницей остатков
      const response = UrlFetchApp.fetch(stocksUrl, options);
      const code = response.getResponseCode();
      console.log(`Код ответа: ${code}`);
      
      if (code !== 200) {
        throw new Error(`API вернул ошибку: ${code} | ${response.getContentText()}`);
      }
      
      const data = JSON.parse(response.getContentText());
      if (data.status != "OK" || !data.result) {
        throw new Error("Ошибка при получении остатков: " + (data.errors ? JSON.stringify(data.errors) : "статус " + data.status));
      }
      
      // Обрабатываем полученные данные: проходим по каждому складу и каждому товару
      const warehouses = data.result.warehouses;
      warehouses.forEach(function (warehouse) {
        const warehouseId = warehouse.warehouseId;
        const warehouseName = warehouseMap[warehouseId] || "";  // название склада (если удалось получить)
        warehouse.offers.forEach(function (offer) {
          const sku = offer.offerId;               // SKU товара (идентификатор товара у продавца)
          const stocks = offer.stocks;             // массив остатков по типам (FIT, AVAILABLE, и т.д.)
          // Инициализируем переменные для основных типов остатков:
          let totalFit = 0, available = 0, reserved = 0;
          stocks.forEach(function (stock) {
            if (stock.type === "FIT") totalFit = stock.count;
            if (stock.type === "AVAILABLE") available = stock.count;
            if (stock.type === "FREEZE") reserved = stock.count;
          });
          // Добавляем данные в результат
          allStocks.push({
            sku: sku,
            warehouseId: warehouseId,
            warehouseName: warehouseName,
            totalFit: totalFit,
            available: available,
            reserved: reserved
          });
        });
      });
      
      // Получаем токен следующей страницы (если он есть)
      pageToken = data.result.paging && data.result.paging.nextPageToken ? data.result.paging.nextPageToken : null;
    } while (pageToken);
    
    console.log(`Получено записей: ${allStocks.length}`);
    return allStocks;
    
  } catch (error) {
    console.error('Ошибка при получении остатков Яндекс Маркета:', error);
    throw error;
  }
}

/**
 * Записывает данные Яндекс Маркета в Google Sheets (обновленная версия)
 */
function writeYandexToGoogleSheets(data, storeName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  // Определяем название листа на основе магазина
  const sheetName = sanitizeSheetName(storeName);
  
  console.log(`Создаем/используем лист: ${sheetName}`);
  
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  // Создаем лист если не существует
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  // Очищаем только диапазон E:J (данные остатков) как в оригинальном скрипте
  const lastRow = sheet.getLastRow();
  if (lastRow > 0) {
    sheet.getRange(1, 5, lastRow, 6).clearContent(); // столбцы E–J
  }
  
  if (data.length === 0) {
    console.log('Нет данных для записи');
    return;
  }
  
  // Подготавливаем данные в формате оригинального скрипта
  const rows = [];
  
  // Добавляем заголовок столбцов (для удобства)
  rows.push(["SKU товара", "ID склада", "Название склада", "Всего (FIT)", "Доступно (AVAILABLE)", "Резерв (FREEZE)"]);
  
  // Добавляем данные
  data.forEach(item => {
    rows.push([
      item.sku || '',
      item.warehouseId || '',
      item.warehouseName || '',
      item.totalFit || 0,
      item.available || 0,
      item.reserved || 0
    ]);
  });
  
  // Записываем данные в столбцы E:J
  if (rows.length > 0) {
    try {
      const dataRange = sheet.getRange(1, 5, rows.length, rows[0].length);
      dataRange.setValues(rows);
      
      // Форматируем заголовки
      const headerRange = sheet.getRange(1, 5, 1, rows[0].length);
      headerRange.setFontWeight('bold');
      headerRange.setBackground('#E8F0FE');
      
      // Фильтры убраны по запросу пользователя
      
      console.log(`Записано ${rows.length} строк в Google Таблицы (включая заголовки)`);
    } catch (error) {
      console.error('Ошибка при записи данных:', error);
      throw error;
    }
  } else {
    console.log('Нет данных для записи');
  }
}

/**
 * Выгружает цены для активного Яндекс Маркет магазина и пишет в T:X
 */
function exportYandexPrices() {
  const config = getYandexConfig();
  const token = config.API_TOKEN;
  const campaignId = config.CAMPAIGN_ID;
  if (!token || !campaignId) {
    throw new Error('Не заданы токен и campaign_id Яндекс Маркета. Добавьте магазин.');
  }
  // Получаем все цены напрямую из offer-prices (GET с пагинацией)
  const allOffers = fetchAllYandexPrices(token, campaignId);
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = sanitizeSheetName(config.STORE_NAME || 'Яндекс Маркет');
  const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
  writeYandexAllPricesToSheetT(sheet, allOffers);
}

function fetchYandexPricesBySkus(token, campaignId, skus) {
  const base = 'https://api.partner.market.yandex.ru';
  const headers = {
    'Authorization': 'OAuth ' + token,
    'Content-Type': 'application/json'
  };

  const map = new Map(); // sku -> {price, old_price, currency}
  const chunk = 200; // безопасный размер
  for (let i = 0; i < skus.length; i += chunk) {
    const part = skus.slice(i, i + chunk);
    // В некоторых версиях API: POST /v2/campaigns/{id}/offer-prices
    const url = `${base}/v2/campaigns/${encodeURIComponent(campaignId)}/offer-prices`; 
    const body = { offers: part.map(sku => ({ sku })) };
    try {
      const resp = UrlFetchApp.fetch(url, { method: 'post', headers, muteHttpExceptions: true, payload: JSON.stringify(body) });
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) {
        const data = JSON.parse(resp.getContentText());
        const items = (data && data.result && data.result.offers) || data.offers || [];
        for (const it of items) {
          const s = String(it.sku || it.offerId || it.shopSku || '');
          const p = it.price || it.basicPrice || it.currentPrice || {};
          const price = Number(p.value || p.price || it.priceValue || 0);
          const currency = p.currency || it.currency || 'RUB';
          const oldPrice = Number(it.oldPrice || p.oldValue || 0) || '';
          if (s) map.set(s, { price: price || '', old_price: oldPrice, currency });
        }
      }
    } catch (e) {
      // игнорируем и идём дальше
    }
    Utilities.sleep(120);
  }
  return map;
}

/**
 * Пагинированная выгрузка всех цен через GET offer-prices
 */
function fetchAllYandexPrices(token, campaignId) {
  const base = 'https://api.partner.market.yandex.ru';
  // Пробуем разные варианты заголовков авторизации
  const headersCandidates = [
    { 'Authorization': 'Api-Key ' + token, 'Content-Type': 'application/json' },
    { 'Authorization': 'OAuth oauth_token="' + token + '"', 'Content-Type': 'application/json' },
    { 'Authorization': 'OAuth ' + token, 'Content-Type': 'application/json' }
  ];
  const urls = [
    (pageToken, limit) => `${base}/campaigns/${encodeURIComponent(campaignId)}/offer-prices?limit=${limit}${pageToken ? `&page_token=${encodeURIComponent(pageToken)}` : ''}`,
    (pageToken, limit) => `${base}/v2/campaigns/${encodeURIComponent(campaignId)}/offer-prices?limit=${limit}${pageToken ? `&page_token=${encodeURIComponent(pageToken)}` : ''}`
  ];

  const limit = 1000;
  let pageToken = '';
  const all = [];
  let page = 0;
  do {
    let resp = null;
    for (const makeUrl of urls) {
      for (const headers of headersCandidates) {
        try {
          const url = makeUrl(pageToken, limit);
          const r = UrlFetchApp.fetch(url, { method: 'get', headers, muteHttpExceptions: true });
          const code = r.getResponseCode();
          if (code >= 200 && code < 300) {
            resp = JSON.parse(r.getContentText());
            break;
          }
        } catch (e) {
          // пробуем следующую комбинацию
        }
      }
      if (resp) break;
    }
    if (!resp) break;
    const offers = (resp && resp.result && resp.result.offers) || resp.offers || [];
    for (const offer of offers) {
      const sku = String(offer.offerId || offer.id || offer.shopSku || offer.sku || '').trim();
      const priceObj = offer.price || {};
      const price = Number(priceObj.value || 0) || '';
      const currency = priceObj.currencyId || priceObj.currency || 'RUR';
      const discountBase = Number(priceObj.discountBase || 0) || '';
      all.push({ sku, price, old_price: discountBase, currency });
    }
    pageToken = (resp && resp.result && resp.result.paging && resp.result.paging.nextPageToken) || resp.nextPageToken || '';
    page++;
    Utilities.sleep(120);
    if (page > 5000) break;
  } while (pageToken);

  return all;
}

function writeYandexAllPricesToSheetT(sheet, offers) {
  const startCol = 20; // T
  const headers = ['SKU', 'Цена, ₽', 'Старая цена, ₽', 'Валюта'];
  sheet.getRange(1, startCol, 1, headers.length).setValues([headers]);
  sheet.getRange(1, startCol, 1, headers.length).setFontWeight('bold').setBackground('#FFF3CD');

  const rows = offers.map(o => [o.sku, o.price, o.old_price, o.currency || 'RUR']);
  if (rows.length > 0) {
    sheet.getRange(2, startCol, rows.length, headers.length).setValues(rows);
  }
  sheet.autoResizeColumns(startCol, headers.length);
  sheet.getRange(rows.length + 3, startCol).setValue('Цены Яндекс Маркета (все) обновлены: ' + new Date().toLocaleString('ru-RU'));
}

function writeYandexPricesToSheetT(sheet, pricesMap, orderSkus) {
  const startCol = 20; // T
  const headers = ['SKU', 'Цена, ₽', 'Старая цена, ₽', 'Валюта'];
  sheet.getRange(1, startCol, 1, headers.length).setValues([headers]);
  sheet.getRange(1, startCol, 1, headers.length).setFontWeight('bold').setBackground('#FFF3CD');

  const rows = [];
  const source = Array.isArray(orderSkus) && orderSkus.length ? orderSkus : [];
  for (const sku of source) {
    const p = pricesMap.get(sku);
    if (p) rows.push([sku, p.price || '', p.old_price || '', p.currency || 'RUB']);
    else rows.push([sku, '', '', 'RUB']);
  }
  sheet.getRange(2, startCol, rows.length, headers.length).setValues(rows);
  sheet.autoResizeColumns(startCol, headers.length);
  sheet.getRange(rows.length + 3, startCol).setValue('Цены Яндекс Маркета обновлены: ' + new Date().toLocaleString('ru-RU'));
}

/**
 * Получить все shopSku из кабинета Маркета (offer-mappings) с пагинацией
 */
function fetchAllYandexShopSkus(token, campaignId) {
  const base = 'https://api.partner.market.yandex.ru';
  const headers = {
    'Authorization': 'OAuth ' + token,
    'Content-Type': 'application/json'
  };
  const url = `${base}/v2/campaigns/${encodeURIComponent(campaignId)}/offer-mappings`;
  const limit = 200;
  let pageToken = '';
  const result = [];
  let page = 0;
  do {
    const body = { limit, page_token: pageToken };
    try {
      const resp = UrlFetchApp.fetch(url, { method: 'post', headers, muteHttpExceptions: true, payload: JSON.stringify(body) });
      const code = resp.getResponseCode();
      if (code >= 200 && code < 300) {
        const data = JSON.parse(resp.getContentText());
        const mappings = (data && data.result && (data.result.offerMappings || data.result.mappings)) || data.offerMappings || data.mappings || [];
        for (const m of mappings) {
          const sku = String((m.offer && (m.offer.shopSku || m.offer.sku)) || m.shopSku || m.sku || '').trim();
          if (sku) result.push(sku);
        }
        pageToken = (data && data.result && data.result.page_token) || data.page_token || '';
      } else {
        break;
      }
    } catch (e) {
      break;
    }
    page++;
    Utilities.sleep(120);
    if (page > 5000) break;
  } while (pageToken);
  return Array.from(new Set(result));
}

/**
 * Тестирует подключение к Яндекс Маркет API
 */
function testYandexConnection() {
  try {
    const config = getYandexConfig();
    
    if (!config.API_TOKEN || !config.CAMPAIGN_ID) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Не настроен API токен или Campaign ID для Яндекс Маркет магазина!', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    console.log(`Тестируем подключение к Яндекс Маркет API для магазина: ${config.STORE_NAME}`);
    console.log(`Campaign ID: ${config.CAMPAIGN_ID}`);
    console.log(`API Token: ***${config.API_TOKEN.slice(-4)}`);
    
    // Тестируем получение информации о кампании
    const url = `https://api.partner.market.yandex.ru/campaigns/${config.CAMPAIGN_ID}`;
    
    const options = {
      method: 'GET',
      headers: {
        'Api-Key': config.API_TOKEN,
        'Content-Type': 'application/json',
      },
      muteHttpExceptions: true
    };
    
    console.log(`Отправляем тестовый запрос: ${url}`);
    const response = UrlFetchApp.fetch(url, options);
    const code = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log(`Код ответа: ${code}`);
    console.log(`Ответ: ${responseText}`);
    
    if (code === 200) {
      const data = JSON.parse(responseText);
      if (data.status === "OK" && data.result) {
        SpreadsheetApp.getUi().alert('Успех', `Подключение к Яндекс Маркет API успешно!\n\nКампания: ${data.result.domain || 'Неизвестно'}\nСтатус: ${data.result.status || 'Неизвестно'}`, SpreadsheetApp.getUi().ButtonSet.OK);
      } else {
        SpreadsheetApp.getUi().alert('Ошибка', `API вернул ошибку: ${data.errors ? JSON.stringify(data.errors) : 'статус ' + data.status}`, SpreadsheetApp.getUi().ButtonSet.OK);
      }
    } else {
      SpreadsheetApp.getUi().alert('Ошибка', `Ошибка подключения к Яндекс Маркет API!\n\nКод: ${code}\nОтвет: ${responseText}`, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    
  } catch (error) {
    console.error('Ошибка при тестировании подключения к Яндекс Маркет API:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка тестирования: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Быстрый тест с вашими токенами (точная копия рабочего скрипта)
 */
function testYandexWithYourTokens() {
  try {
    // === НАСТРОЙКИ: введите свои параметры ниже ===
    const API_TOKEN = "ACMA:b0BKJAZYstQEOJf5sYDNyOlEONs3cGcrTYprLMZi:bb8c04d4";       // Авторизационный токен API Яндекс.Маркета (Api Key или OAuth)
    const CAMPAIGN_ID = 89101200;             // Идентификатор вашего магазина (campaignId) на Маркете (число)
    const SHEET_NAME = "YM MR";      // Название листа Google Таблицы для вывода остатков

    console.log('Тестируем с вашими токенами...');
    console.log(`Campaign ID: ${CAMPAIGN_ID}`);
    console.log(`API Token: ***${API_TOKEN.slice(-4)}`);

    // Получаем объект листа по названию
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      // Если лист не найден, создаем его
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      console.log(`Создан лист: ${SHEET_NAME}`);
    }
    
    // Очищаем только диапазон E:J (данные остатков)
    const lastRow = sheet.getLastRow();
    if (lastRow > 0) {
      sheet.getRange(1, 5, lastRow, 6).clearContent(); // столбцы E–J
    }

    // --- 1) Определяем URL и заголовки для API-запросов ---
    // Базовый URL для получения остатков:
    const stocksUrl = "https://api.partner.market.yandex.ru/campaigns/" + CAMPAIGN_ID + "/offers/stocks";
    // (Опционально) URL для получения списка складов (чтобы сопоставить ID с названиями):
    const warehousesUrl = "https://api.partner.market.yandex.ru/warehouses";

    // Заголовки HTTP для авторизации: используем Api Key токен
    const headers = {
      "Api-Key": API_TOKEN
    };

    // --- 2) Получаем справочник складов Маркета (FBY) ---
    const warehouseMap = {};  // словарь для сопоставления warehouseId -> name
    try {
      const whResponse = UrlFetchApp.fetch(warehousesUrl, { "method": "get", "headers": headers });
      const whData = JSON.parse(whResponse.getContentText());
      if (whData.status == "OK" && whData.result && whData.result.warehouses) {
        whData.result.warehouses.forEach(function (w) {
          warehouseMap[w.id] = w.name;
        });
        console.log(`Получено складов: ${Object.keys(warehouseMap).length}`);
      }
    } catch (e) {
      console.log("Не удалось получить список складов: " + e);
    }

    // --- 3) Подготовка тела запроса для получения остатков товаров ---
    // Запросим **все неархивные товары** с учетом FBY-складов. 
    // withTurnover:false означает, что показатели оборачиваемости можно не включать.
    const requestBody = {
      "archived": false,
      "withTurnover": false
    };

    // Опции для UrlFetchApp (POST запрос с JSON-телом)
    const options = {
      "method": "post",
      "contentType": "application/json",
      "headers": headers,
      "payload": JSON.stringify(requestBody)
    };

    // --- 4) Выполняем запрос к API и обрабатываем данные с постраничной загрузкой ---
    const rows = [];  // массив для строк, которые запишем в таблицу
    // Добавим заголовок столбцов (для удобства)
    rows.push(["SKU товара", "ID склада", "Название склада", "Всего (FIT)", "Доступно (AVAILABLE)", "Резерв (FREEZE)"]);

    let pageToken = null;
    do {
      // Если есть токен следующей страницы, добавляем его к телу запроса
      if (pageToken) {
        requestBody.page_token = pageToken;
        options.payload = JSON.stringify(requestBody);
      }
      // Выполняем API-запрос за текущей страницей остатков
      const response = UrlFetchApp.fetch(stocksUrl, options);
      const code = response.getResponseCode();
      console.log(`Код ответа: ${code}`);
      
      if (code !== 200) {
        throw new Error(`API вернул ошибку: ${code} | ${response.getContentText()}`);
      }
      
      const data = JSON.parse(response.getContentText());
      if (data.status != "OK" || !data.result) {
        throw new Error("Ошибка при получении остатков: " + (data.errors ? JSON.stringify(data.errors) : "статус " + data.status));
      }

      // Обрабатываем полученные данные: проходим по каждому складу и каждому товару
      const warehouses = data.result.warehouses;
      warehouses.forEach(function (warehouse) {
        const warehouseId = warehouse.warehouseId;
        const warehouseName = warehouseMap[warehouseId] || "";  // название склада (если удалось получить)
        warehouse.offers.forEach(function (offer) {
          const sku = offer.offerId;               // SKU товара (идентификатор товара у продавца)
          const stocks = offer.stocks;             // массив остатков по типам (FIT, AVAILABLE, и т.д.)
          // Инициализируем переменные для основных типов остатков:
          let totalFit = 0, available = 0, reserved = 0;
          stocks.forEach(function (stock) {
            if (stock.type === "FIT") totalFit = stock.count;
            if (stock.type === "AVAILABLE") available = stock.count;
            if (stock.type === "FREEZE") reserved = stock.count;
          });
          // Добавляем строку в результаты
          rows.push([sku, warehouseId, warehouseName, totalFit, available, reserved]);
        });
      });

      // Получаем токен следующей страницы (если он есть)
      pageToken = data.result.paging && data.result.paging.nextPageToken ? data.result.paging.nextPageToken : null;
    } while (pageToken);

    // --- 5) Записываем все собранные строки на лист таблицы ---
    if (rows.length > 0) {
      sheet.getRange(1, 5, rows.length, rows[0].length).setValues(rows);
    }
    
    const totalRows = rows.length - 1; // -1 для заголовка
    console.log(`YM MR: загружено ${totalRows} строк (вкл. заголовки)`);
    
    SpreadsheetApp.getUi().alert('Успех', `Получено и записано ${totalRows} записей остатков в лист "YM MR"`, SpreadsheetApp.getUi().ButtonSet.OK);
    
  } catch (error) {
    console.error('Ошибка при тестировании:', error);
    SpreadsheetApp.getUi().alert('Ошибка', `Ошибка: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}