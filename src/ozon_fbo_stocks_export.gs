/**
 * Скрипт для выгрузки остатков со складов FBO Ozon Seller в Google Таблицы
 * Требует настройки API ключей Ozon и Google Sheets
 */

// Конфигурация API Ozon
const OZON_CONFIG = {
  CLIENT_ID: 'YOUR_CLIENT_ID', // Замените на ваш Client ID
  API_KEY: 'YOUR_API_KEY',     // Замените на ваш API ключ
  BASE_URL: 'https://api-seller.ozon.ru'
};

// ID Google Таблицы (из URL)
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID'; // Замените на ID вашей таблицы

/**
 * Основная функция для запуска выгрузки
 */
function exportFBOStocks() {
  try {
    console.log('Начинаем выгрузку остатков FBO...');
    
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
  const url = `${OZON_CONFIG.BASE_URL}/v1/warehouse/list`;
  
  const options = {
    method: 'POST',
    headers: {
      'Client-Id': OZON_CONFIG.CLIENT_ID,
      'Api-Key': OZON_CONFIG.API_KEY,
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
  // Пробуем разные версии API
  const apiVersions = ['/v2/product/info/stocks', '/v1/product/info/stocks', '/v3/product/info/stocks'];
  
  for (let i = 0; i < apiVersions.length; i++) {
    try {
      const url = `${OZON_CONFIG.BASE_URL}${apiVersions[i]}`;
      
      const options = {
        method: 'POST',
        headers: {
          'Client-Id': OZON_CONFIG.CLIENT_ID,
          'Api-Key': OZON_CONFIG.API_KEY,
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
  const spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
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
