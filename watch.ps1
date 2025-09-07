# Скрипт для автоматического отслеживания изменений и отправки в Google Apps Script
# Аналог clasp watch для новых версий

Write-Host "Запуск автоматического отслеживания изменений..." -ForegroundColor Green
Write-Host "Нажмите Ctrl+C для остановки" -ForegroundColor Yellow

$lastWriteTime = @{}

# Функция для отправки изменений
function Push-Changes {
    Write-Host "`nОбнаружены изменения, отправляем в Google Apps Script..." -ForegroundColor Yellow
    clasp push
    if ($LASTEXITCODE -eq 0) {
        Write-Host "✅ Изменения успешно отправлены!" -ForegroundColor Green
    } else {
        Write-Host "❌ Ошибка при отправке изменений" -ForegroundColor Red
    }
}

# Отслеживание изменений в папке src
while ($true) {
    $files = Get-ChildItem -Path "src" -Recurse -File | Where-Object { $_.Extension -match '\.(gs|js|json)$' }
    
    foreach ($file in $files) {
        $currentWriteTime = $file.LastWriteTime
        if ($lastWriteTime.ContainsKey($file.FullName)) {
            if ($currentWriteTime -gt $lastWriteTime[$file.FullName]) {
                $lastWriteTime[$file.FullName] = $currentWriteTime
                Push-Changes
                break
            }
        } else {
            $lastWriteTime[$file.FullName] = $currentWriteTime
        }
    }
    
    Start-Sleep -Seconds 2
}
