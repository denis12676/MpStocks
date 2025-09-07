# Настройка автоматической отправки в Google Apps Script

## 1. Установка Google Apps Script CLI

```powershell
# Установка Node.js (если не установлен)
# Скачайте с https://nodejs.org/

# Установка @google/clasp
npm install -g @google/clasp
```

## 2. Авторизация в Google

```powershell
# Авторизация в Google аккаунте
clasp login
```

## 3. Автоматическая отправка

### Вариант 1: Через PowerShell скрипт
```powershell
# Запуск автоматического деплоя
.\deploy.ps1
```

### Вариант 2: Через npm команды
```powershell
# Отправка кода
npm run push

# Открытие проекта в браузере
npm run open

# Просмотр логов
npm run logs
```

### Вариант 3: Прямые команды clasp
```powershell
# Отправка кода
clasp push

# Открытие проекта
clasp open

# Создание деплоя
clasp deploy
```

## 4. Структура проекта

```
MpStocks/
├── src/
│   └── ozon_fbo_stocks_export.gs  # Основной скрипт
├── .clasp.json                    # Конфигурация проекта
├── .claspignore                   # Игнорируемые файлы
├── package.json                   # npm конфигурация
├── deploy.ps1                     # Скрипт автоматического деплоя
└── README_ozon_setup.md          # Инструкция по настройке
```

## 5. Настройка API ключей

В файле `src/ozon_fbo_stocks_export.gs` замените:
- `YOUR_CLIENT_ID` на ваш Client ID от Ozon
- `YOUR_API_KEY` на ваш API ключ от Ozon  
- `YOUR_SPREADSHEET_ID` на ID вашей Google Таблицы

## 6. Автоматизация

Для автоматической отправки при каждом коммите добавьте в `.git/hooks/post-commit`:

```bash
#!/bin/sh
powershell -ExecutionPolicy Bypass -File deploy.ps1
```

## Команды для быстрого деплоя

```powershell
# Полная настройка (выполнить один раз)
npm install -g @google/clasp
clasp login

# Ежедневное использование
.\deploy.ps1
```
