/***** CONFIG *****/
// Укажи email для оповещений об ошибках (можно оставить пустым)
const ALERT_EMAIL = ''; // например: 'you@example.com'
// Вставь webhook URL своего Telegram-бота (опционально)
const TELEGRAM_WEBHOOK = ''; // например: 'https://api.telegram.org/bot<token>/sendMessage?chat_id=<id>'

/***** ENTRYPOINT *****/
// Основной запуск по триггеру (раз в час)
function syncAllStocksHourly() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.warn('Запуск пропущен: прошлый ещё выполняется.');
    return;
  }

  const started = new Date();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('stocks.lastRunStartedAt', started.toISOString());

  const results = [];
  try {
    results.push(runStep('WB', exportAllWBStoresStocksStatisticsAPI));
    results.push(runStep('YM', exportAllYandexStoresStocks));
    results.push(runStep('OZON', exportAllStoresStocks));
  } finally {
    lock.releaseLock();
    props.setProperty('stocks.lastRunFinishedAt', new Date().toISOString());
  }

  const durationSec = Math.round((Date.now() - started.getTime()) / 1000);
  const failed = results.filter(r => !r.ok);
  console.log('syncAll summary:', JSON.stringify({ durationSec, results }, null, 2));

  if (failed.length) {
    const msg = [
      `Ошибки при выгрузке остатков ( ${durationSec}s ):`,
      ...failed.map(f => `• ${f.name}: ${f.error}`)
    ].join('\n');
    notify(msg);
  }
}

/***** STEP WRAPPER *****/
function runStep(name, fn) {
  try {
    const data = runWithRetry(() => {
      // если твои функции что-то возвращают — попадёт в data
      return fn();
    }, { tries: 5, baseMs: 600, maxMs: 8000 });
    return { name, ok: true, data: toPlainJsonSafe(data) };
  } catch (e) {
    const msg = e && e.stack ? e.stack : String(e);
    console.error(`${name} failed:`, msg);
    return { name, ok: false, error: String(e) };
  }
}

function toPlainJsonSafe(v) {
  try { return JSON.parse(JSON.stringify(v)); } catch (_) { return String(v); }
}

/***** RETRY WITH EXPONENTIAL BACKOFF *****/
function runWithRetry(fn, opts) {
  const tries = opts?.tries ?? 4;
  const baseMs = opts?.baseMs ?? 400;
  const maxMs  = opts?.maxMs  ?? 7000;

  let lastErr;
  for (let i = 0; i < tries; i++) {
    try {
      return fn();
    } catch (e) {
      lastErr = e;
      const msg = String(e);
      // Повторяем только на типовых временных сбоях API/сети/квотах:
      const retryable = /429|Rate|Too Many|Timeout|Timed out|Exceeded|Quota|Internal|Network|Service|5\d\d/i.test(msg);
      if (!retryable || i === tries - 1) throw e;

      const delay = Math.min(maxMs, baseMs * Math.pow(2, i)) + Math.floor(Math.random() * 250);
      console.warn(`Retry ${i + 1}/${tries} через ${delay} мс из-за: ${msg}`);
      Utilities.sleep(delay);
    }
  }
  throw lastErr;
}

/***** NOTIFY *****/
function notify(text) {
  try {
    if (ALERT_EMAIL) MailApp.sendEmail(ALERT_EMAIL, 'Stocks sync: ошибки', text);
  } catch (e) { console.warn('Mail notify failed', e); }
  try {
    if (TELEGRAM_WEBHOOK) {
      UrlFetchApp.fetch(TELEGRAM_WEBHOOK, {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ text })
      });
    }
  } catch (e) { console.warn('TG notify failed', e); }
}

/***** TRIGGERS *****/
// Создать/обновить единый почасовой триггер
function createHourlyTrigger() {
  // Чистим дубликаты
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncAllStocksHourly') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('syncAllStocksHourly')
    .timeBased()
    .everyHours(1)   // раз в час
    .create();

  console.log('Триггер создан: раз в час -> syncAllStocksHourly');
}

// На всякий случай — функция удаления ВСЕХ триггеров этого проекта
function deleteAllTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  console.log('Все триггеры удалены.');
}

/*** PRICES: hourly sync for Ozon + WB ***/
function syncAllPricesHourly() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    console.warn('Запуск цен пропущен: прошлый ещё выполняется.');
    return;
  }

  const started = new Date();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('prices.lastRunStartedAt', started.toISOString());

  const results = [];
  try {
    // Ozon: детальные цены по всем магазинам
    results.push(runStep('OZON_PRICES', exportAllStoresPricesDetailed));
    // Wildberries: цены по всем магазинам
    results.push(runStep('WB_PRICES', exportAllWBStoresPrices));
  } finally {
    lock.releaseLock();
    props.setProperty('prices.lastRunFinishedAt', new Date().toISOString());
  }

  const durationSec = Math.round((Date.now() - started.getTime()) / 1000);
  const failed = results.filter(r => !r.ok);
  console.log('syncAllPrices summary:', JSON.stringify({ durationSec, results }, null, 2));

  if (failed.length) {
    const msg = [
      `Ошибки при выгрузке цен ( ${durationSec}s ):`,
      ...failed.map(f => `• ${f.name}: ${f.error}`)
    ].join('\n');
    notify(msg);
  }
}

// Создать/обновить почасовой триггер для цен
function createHourlyPricesTrigger() {
  // Чистим дубликаты этого обработчика
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'syncAllPricesHourly') ScriptApp.deleteTrigger(t);
  });

  ScriptApp.newTrigger('syncAllPricesHourly')
    .timeBased()
    .everyHours(1)
    .create();

  console.log('Триггер создан: раз в час -> syncAllPricesHourly');
}

// Удобный хелпер: создать оба триггера (остатки и цены)
function createAllHourlyTriggers() {
  createHourlyTrigger();
  createHourlyPricesTrigger();
}