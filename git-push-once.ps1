# Однократная отправка изменений в Git

Write-Host "Отправка изменений в Git..." -ForegroundColor Green

# Проверяем статус
Write-Host "Проверяем статус репозитория..." -ForegroundColor Yellow
git status

# Добавляем все изменения
Write-Host "`nДобавляем все изменения..." -ForegroundColor Yellow
git add .

# Коммитим с временной меткой
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
$commitMessage = "Update: $timestamp"
Write-Host "Коммитим изменения: $commitMessage" -ForegroundColor Yellow
git commit -m $commitMessage

if ($LASTEXITCODE -eq 0) {
    # Пушим в репозиторий
    Write-Host "`nОтправляем в GitHub..." -ForegroundColor Yellow
    git push
    
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Изменения успешно отправлены в Git!" -ForegroundColor Green
    } else {
        Write-Host "❌ Ошибка при отправке в Git" -ForegroundColor Red
    }
} else {
    Write-Host "❌ Ошибка при коммите" -ForegroundColor Red
}

Write-Host "`nГотово!" -ForegroundColor Green
