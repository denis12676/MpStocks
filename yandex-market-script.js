// ================================
// СКРИПТ ДЛЯ ВЫГРУЗКИ ЦЕН ИЗ ЯНДЕКС МАРКЕТА В GOOGLE SHEETS
// ================================

// ================================
// НАСТРОЙКИ - ИЗМЕНИТЕ НА СВОИ!
// ================================

// Ваш токен от Яндекс Маркета
const API_TOKEN = 'AQAAAAAAAA...'; // Замените на ваш токен

// ID кампании (магазина) - можно узнать через getCampaigns()
const CAMPAIGN_ID = '12345678'; // Замените на ваш campaign ID

// Тип авторизации: true = API-Key токен, false = OAuth токен
const USE_API_KEY = true;

// OAuth параметры (если используете OAuth токен)
const OAUTH_CLIENT_ID = 'ваш_client_id'; // Нужно только для OAuth

// ================================
// ОСНОВНЫЕ ФУНКЦИИ
// ================================

/**
 * Главная функция для выгрузки всех цен
 */
function exportAllPrices() {
  try {
    console.log('Начинаем выгрузку цен...');
    
    const sheet = getOrCreateSheet('Цены товаров');
    clearAndSetupSheet(sheet);
    
    let pageToken = null;
    let totalProcessed = 0;
    let rowIndex = 2; // Начинаем с второй строки (после заголовков)
    
    do {
      const response = fetchPricesFromAPI(pageToken);
      
      if (response.result && response.result.offers) {
        const offers = response.result.offers;
        console.log(`Получено ${offers.length} товаров`);
        
        // Записываем данные в таблицу
        const data = offers.map(offer => [
          offer.id || offer.offerId || '',
          offer.marketSku || '',
          offer.price ? offer.price.value : '',
          offer.price ? offer.price.currencyId : 'RUR',
          offer.price ? offer.price.discountBase || '' : '',
          offer.price ? offer.price.vat || '' : '',
          offer.updatedAt ? formatDate(offer.updatedAt) : ''
        ]);
        
        // Записываем блок данных
        if (data.length > 0) {
          const range = sheet.getRange(rowIndex, 1, data.length, 7);
          range.setValues(data);
          rowIndex += data.length;
        }
        
        totalProcessed += offers.length;
        pageToken = response.result.paging ? response.result.paging.nextPageToken : null;
        
        // Добавляем небольшую задержку для соблюдения лимитов API
        if (pageToken) {
          Utilities.sleep(100);
        }
        
      } else {
        console.error('Неожиданный формат ответа:', response);
        break;
      }
    } while (pageToken);
    
    // Добавляем информацию об обновлении
    sheet.getRange(1, 9).setValue('Обновлено:');
    sheet.getRange(1, 10).setValue(new Date());
    sheet.getRange(2, 9).setValue('Всего товаров:');
    sheet.getRange(2, 10).setValue(totalProcessed);
    
    console.log(`✅ Успешно выгружено ${totalProcessed} товаров`);
    SpreadsheetApp.getUi().alert(`Выгрузка завершена!\n\nВыгружено товаров: ${totalProcessed}\nЛист: "${sheet.getName()}"`);
    
  } catch (error) {
    console.error('Ошибка при выгрузке:', error);
    SpreadsheetApp.getUi().alert(`Ошибка: ${error.message}`);
  }
}

/**
 * Выгрузка цен конкретных товаров по их SKU
 */
function exportSpecificPrices(offerIds = []) {
  // Если массив пустой, берем SKU из диапазона A2:A в текущем листе
  if (offerIds.length === 0) {
    const sheet = SpreadsheetApp.getActiveSheet();
    const values = sheet.getRange('A2:A').getValues().filter(row => row[0] !== '');
    offerIds = values.map(row => row[0].toString());
  }
  
  if (offerIds.length === 0) {
    SpreadsheetApp.getUi().alert('Не найдены SKU для выгрузки. Укажите их в столбце A или передайте в параметрах функции.');
    return;
  }
  
  try {
    const sheet = getOrCreateSheet('Конкретные цены');
    clearAndSetupSheet(sheet);
    
    // API позволяет запрашивать до 500 товаров за раз
    const batchSize = 500;
    let totalProcessed = 0;
    let rowIndex = 2;
    
    for (let i = 0; i < offerIds.length; i += batchSize) {
      const batch = offerIds.slice(i, i + batchSize);
      const response = fetchSpecificPrices(batch);
      
      if (response.result && response.result.offers) {
        const offers = response.result.offers;
        
        const data = offers.map(offer => [
          offer.offerId || offer.id || '',
          offer.marketSku || '',
          offer.price ? offer.price.value : '',
          offer.price ? offer.price.currencyId : 'RUR',
          offer.price ? offer.price.discountBase || '' : '',
          offer.price ? offer.price.vat || '' : '',
          offer.updatedAt ? formatDate(offer.updatedAt) : ''
        ]);
        
        if (data.length > 0) {
          const range = sheet.getRange(rowIndex, 1, data.length, 7);
          range.setValues(data);
          rowIndex += data.length;
        }
        
        totalProcessed += offers.length;
      }
      
      // Задержка между запросами
      if (i + batchSize < offerIds.length) {
        Utilities.sleep(200);
      }
    }
    
    sheet.getRange(1, 9).setValue('Обновлено:');
    sheet.getRange(1, 10).setValue(new Date());
    sheet.getRange(2, 9).setValue('Найдено товаров:');
    sheet.getRange(2, 10).setValue(totalProcessed);
    
    console.log(`✅ Успешно выгружено ${totalProcessed} из ${offerIds.length} запрошенных товаров`);
    SpreadsheetApp.getUi().alert(`Выгрузка завершена!\n\nНайдено товаров: ${totalProcessed} из ${offerIds.length} запрошенных`);
    
  } catch (error) {
    console.error('Ошибка при выгрузке конкретных товаров:', error);
    SpreadsheetApp.getUi().alert(`Ошибка: ${error.message}`);
  }
}

/**
 * Получение списка всех кампаний и магазинов
 */
function getCampaigns() {
  try {
    const url = 'https://api.partner.market.yandex.ru/campaigns';
    const options = {
      method: 'GET',
      headers: getAuthHeaders(),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseData = JSON.parse(response.getContentText());
    
    if (responseCode !== 200) {
      throw new Error(`API вернул код ${responseCode}: ${responseData.message || 'Unknown error'}`);
    }
    
    const sheet = getOrCreateSheet('Список кампаний');
    sheet.clear();
    
    // Заголовки
    sheet.getRange(1, 1, 1, 5).setValues([['Campaign ID', 'Название', 'Business ID', 'Бизнес', 'Модель размещения']]);
    sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#e1f5fe');
    
    if (responseData.campaigns && responseData.campaigns.length > 0) {
      const data = responseData.campaigns.map(campaign => [
        campaign.id,
        campaign.domain || '',
        campaign.business ? campaign.business.id : '',
        campaign.business ? campaign.business.name : '',
        campaign.placementType || ''
      ]);
      
      sheet.getRange(2, 1, data.length, 5).setValues(data);
      
      // Автоподбор ширины столбцов
      for (let i = 1; i <= 5; i++) {
        sheet.autoResizeColumn(i);
      }
      
      console.log(`Найдено кампаний: ${responseData.campaigns.length}`);
      SpreadsheetApp.getUi().alert(`Найдено ${responseData.campaigns.length} кампаний.\n\nИспользуйте значения из колонки "Campaign ID" в настройках скрипта.`);
    } else {
      SpreadsheetApp.getUi().alert('Кампании не найдены. Проверьте права доступа токена.');
    }
    
  } catch (error) {
    console.error('Ошибка получения кампаний:', error);
    SpreadsheetApp.getUi().alert(`Ошибка получения списка кампаний: ${error.message}`);
  }
}

// ================================
// ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
// ================================

/**
 * Формирование заголовков авторизации
 */
function getAuthHeaders() {
  if (USE_API_KEY) {
    return {
      'Authorization': `Api-Key ${API_TOKEN}`,
      'Content-Type': 'application/json'
    };
  } else {
    return {
      'Authorization': `OAuth oauth_token="${API_TOKEN}", oauth_client_id="${OAUTH_CLIENT_ID}"`,
      'Content-Type': 'application/json'
    };
  }
}

/**
 * Запрос к API для получения всех цен
 */
function fetchPricesFromAPI(pageToken = null, limit = 1000) {
  let url = `https://api.partner.market.yandex.ru/campaigns/${CAMPAIGN_ID}/offer-prices`;
  
  const params = [];
  if (limit) params.push(`limit=${limit}`);
  if (pageToken) params.push(`page_token=${encodeURIComponent(pageToken)}`);
  
  if (params.length > 0) {
    url += '?' + params.join('&');
  }
  
  const options = {
    method: 'GET',
    headers: getAuthHeaders(),
    muteHttpExceptions: true
  };
  
  console.log(`Запрос: ${url}`);
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseData = JSON.parse(response.getContentText());
  
  if (responseCode !== 200) {
    console.error('Ошибка API:', responseData);
    throw new Error(`API вернул код ${responseCode}: ${responseData.message || responseData.errors?.[0]?.message || 'Unknown error'}`);
  }
  
  return responseData;
}

/**
 * Запрос цен конкретных товаров
 */
function fetchSpecificPrices(offerIds) {
  const url = `https://api.partner.market.yandex.ru/campaigns/${CAMPAIGN_ID}/offer-prices`;
  
  const payload = {
    offerIds: offerIds
  };
  
  const options = {
    method: 'POST',
    headers: getAuthHeaders(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  
  console.log(`Запрос цен для ${offerIds.length} товаров`);
  
  const response = UrlFetchApp.fetch(url, options);
  const responseCode = response.getResponseCode();
  const responseData = JSON.parse(response.getContentText());
  
  if (responseCode !== 200) {
    console.error('Ошибка API:', responseData);
    throw new Error(`API вернул код ${responseCode}: ${responseData.message || responseData.errors?.[0]?.message || 'Unknown error'}`);
  }
  
  return responseData;
}

/**
 * Получение или создание листа
 */
function getOrCreateSheet(sheetName) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * Очистка и настройка листа
 */
function clearAndSetupSheet(sheet) {
  sheet.clear();
  
  // Заголовки
  const headers = [
    'SKU товара',
    'Market SKU',
    'Цена',
    'Валюта',
    'Цена без скидки',
    'НДС',
    'Дата обновления'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#e8f5e8');
  
  // Автоподбор ширины
  for (let i = 1; i <= headers.length; i++) {
    sheet.autoResizeColumn(i);
  }
  
  // Заморозка заголовка
  sheet.setFrozenRows(1);
}

/**
 * Форматирование даты
 */
function formatDate(dateString) {
  try {
    return new Date(dateString).toLocaleString('ru-RU');
  } catch (e) {
    return dateString;
  }
}

// ================================
// ДОПОЛНИТЕЛЬНЫЕ ФУНКЦИИ
// ================================

/**
 * Тестовая функция для проверки подключения
 */
function testConnection() {
  try {
    console.log('Тестируем подключение к API...');
    
    const url = 'https://api.partner.market.yandex.ru/campaigns';
    const options = {
      method: 'GET',
      headers: getAuthHeaders(),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode === 200) {
      console.log('✅ Подключение к API успешно!');
      SpreadsheetApp.getUi().alert('✅ Подключение к API Яндекс Маркета успешно!\n\nМожно приступать к выгрузке данных.');
    } else {
      const errorData = JSON.parse(response.getContentText());
      console.error('❌ Ошибка подключения:', errorData);
      SpreadsheetApp.getUi().alert(`❌ Ошибка подключения к API:\n\nКод: ${responseCode}\nСообщение: ${errorData.message || 'Unknown error'}\n\nПроверьте токен и настройки.`);
    }
    
  } catch (error) {
    console.error('❌ Ошибка при тестировании:', error);
    SpreadsheetApp.getUi().alert(`❌ Ошибка при тестировании подключения:\n\n${error.message}`);
  }
}

/**
 * Создание меню в Google Sheets
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Яндекс Маркет')
    .addItem('📋 Получить список кампаний', 'getCampaigns')
    .addItem('🔄 Тестировать подключение', 'testConnection')
    .addSeparator()
    .addItem('📊 Выгрузить все цены', 'exportAllPrices')
    .addItem('🎯 Выгрузить конкретные цены', 'exportSpecificPrices')
    .addToUi();
}