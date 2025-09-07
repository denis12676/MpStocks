# Скрипт для автоматической отправки в Google Apps Script
# Требует установки @google/clasp: npm install -g @google/clasp

Write-Host "Начинаем деплой в Google Apps Script..." -ForegroundColor Green

# Проверяем установку clasp
try {
    $claspVersion = clasp --version
    Write-Host "Clasp версия: $claspVersion" -ForegroundColor Yellow
} catch {
    Write-Host "Ошибка: @google/clasp не установлен!" -ForegroundColor Red
    Write-Host "Установите командой: npm install -g @google/clasp" -ForegroundColor Yellow
    exit 1
}

# Проверяем авторизацию
try {
    $loginStatus = clasp whoami
    if ($loginStatus -match "Not logged in" -or $LASTEXITCODE -ne 0) {
        Write-Host "Требуется авторизация в Google..." -ForegroundColor Yellow
        clasp login
    } else {
        Write-Host "Авторизован как: $loginStatus" -ForegroundColor Green
    }
} catch {
    Write-Host "Ошибка проверки авторизации" -ForegroundColor Red
    exit 1
}

# Отправляем код в Google Apps Script
Write-Host "Отправляем код в Google Apps Script..." -ForegroundColor Green
clasp push

if ($LASTEXITCODE -eq 0) {
    Write-Host "✅ Код успешно отправлен в Google Apps Script!" -ForegroundColor Green
    
    # Открываем проект в браузере
    Write-Host "Открываем проект в браузере..." -ForegroundColor Yellow
    clasp open
} else {
    Write-Host "❌ Ошибка при отправке кода" -ForegroundColor Red
    exit 1
}

Write-Host "Деплой завершен!" -ForegroundColor Green
