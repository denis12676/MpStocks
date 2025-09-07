# Автоматическая отправка изменений в Git
# Отслеживает изменения и автоматически коммитит и пушит

Write-Host "Запуск автоматической отправки в Git..." -ForegroundColor Green
Write-Host "Нажмите Ctrl+C для остановки" -ForegroundColor Yellow

$lastWriteTime = @{}
$commitCount = 0

# Функция для отправки изменений в Git
function Push-ToGit {
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $commitCount++
    
    Write-Host "`nОбнаружены изменения, отправляем в Git..." -ForegroundColor Yellow
    
    # Добавляем все изменения
    git add .
    
    # Коммитим с временной меткой
    $commitMessage = "Auto-commit #$commitCount - $timestamp"
    git commit -m $commitMessage
    
    if ($LASTEXITCODE -eq 0) {
        # Пушим в репозиторий
        git push
        
        if ($LASTEXITCODE -eq 0) {
            Write-Host "✅ Изменения успешно отправлены в Git! (Коммит #$commitCount)" -ForegroundColor Green
        } else {
            Write-Host "❌ Ошибка при отправке в Git" -ForegroundColor Red
        }
    } else {
        Write-Host "❌ Ошибка при коммите" -ForegroundColor Red
    }
}

# Отслеживание изменений во всех файлах проекта
while ($true) {
    $files = Get-ChildItem -Path "." -Recurse -File | Where-Object { 
        $_.Name -notmatch '\.(git|log)$' -and 
        $_.FullName -notmatch '\.git\\' 
    }
    
    $hasChanges = $false
    
    foreach ($file in $files) {
        $currentWriteTime = $file.LastWriteTime
        if ($lastWriteTime.ContainsKey($file.FullName)) {
            if ($currentWriteTime -gt $lastWriteTime[$file.FullName]) {
                $lastWriteTime[$file.FullName] = $currentWriteTime
                $hasChanges = $true
                break
            }
        } else {
            $lastWriteTime[$file.FullName] = $currentWriteTime
        }
    }
    
    if ($hasChanges) {
        Push-ToGit
    }
    
    Start-Sleep -Seconds 5
}
