#Requires -Version 5.1
<#
.SYNOPSIS
    TechnoPan — установка: проверяет наличие exe и создаёт ярлык на рабочем столе.
    Запускается через setup.bat (двойной клик) из папки дистрибутива.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$AppName         = "TechnoPan"
$ExeName         = "TechnoPanSpec.exe"
$ScriptRoot      = Split-Path -Parent $MyInvocation.MyCommand.Definition
$ExePath         = Join-Path $ScriptRoot $ExeName
$DesktopShortcut = Join-Path ([Environment]::GetFolderPath('Desktop')) "$AppName.lnk"
$StartMenuDir    = Join-Path ([Environment]::GetFolderPath('StartMenu')) "Programs\$AppName"
$StartMenuLink   = Join-Path $StartMenuDir "$AppName.lnk"

function Write-Step { param($msg) Write-Host "`n>>> $msg" -ForegroundColor Cyan }
function Write-OK   { param($msg) Write-Host "    [OK] $msg" -ForegroundColor Green }
function Write-Warn { param($msg) Write-Host "    [!]  $msg" -ForegroundColor Yellow }
function Write-Fail { param($msg) Write-Host "    [X]  $msg" -ForegroundColor Red }
function Write-Info { param($msg) Write-Host "         $msg" -ForegroundColor Gray }

Write-Host ""
Write-Host "============================================================" -ForegroundColor White
Write-Host "  $AppName — Установка" -ForegroundColor White
Write-Host "============================================================" -ForegroundColor White

# ── 1. Проверяем exe ──────────────────────────────────────────────────────────
Write-Step "Проверка файлов программы..."

if (-not (Test-Path $ExePath)) {
    Write-Fail "Не найден $ExeName в папке:`n         $ScriptRoot"
    Write-Info "Убедитесь, что setup.bat лежит рядом с $ExeName."
    exit 1
}
Write-OK "Исполняемый файл: $ExePath"

# Проверяем папку configs
$ConfigsDir = Join-Path $ScriptRoot "configs"
if (Test-Path $ConfigsDir) {
    $cfgCount = (Get-ChildItem $ConfigsDir -Filter "*.yml").Count
    Write-OK "Конфиги: $cfgCount файл(ов) в папке configs\"
} else {
    Write-Warn "Папка configs\ не найдена — программа запустится, но без конфигураций"
}

# ── 2. Проверяем внутренние файлы PyInstaller ─────────────────────────────────
Write-Step "Проверка целостности дистрибутива..."

$InternalDir = Join-Path $ScriptRoot "_internal"
if (Test-Path $InternalDir) {
    Write-OK "Папка _internal\ присутствует"
} else {
    Write-Warn "Папка _internal\ не найдена — дистрибутив может быть неполным"
}

# ── 3. Проверяем, что exe вообще запускается ──────────────────────────────────
Write-Step "Проверка запуска программы..."

try {
    # Запускаем с таймаутом 5 сек; GUI само закроется при --version или просто выйдет
    $proc = Start-Process -FilePath $ExePath -ArgumentList '--help' `
        -PassThru -WindowStyle Hidden -ErrorAction Stop
    $proc.WaitForExit(5000) | Out-Null
    # Если процесс ещё жив (GUI открылось) — убиваем, это нормально
    if (-not $proc.HasExited) { $proc.Kill() }
    Write-OK "Программа запустилась успешно"
} catch {
    Write-Warn "Не удалось выполнить тестовый запуск: $_"
    Write-Info "Попробуйте запустить $ExeName вручную"
}

# ── 4. ODA File Converter ─────────────────────────────────────────────────────
Write-Step "Проверка ODA File Converter (нужен для открытия .dwg)..."

# Повторяем ту же логику поиска, что в odafc_utils.py
function Find-OdaExe {
    $envPath = [Environment]::GetEnvironmentVariable('ODA_FILE_CONVERTER_EXE', 'User')
    if (-not $envPath) {
        $envPath = [Environment]::GetEnvironmentVariable('ODA_FILE_CONVERTER_EXE', 'Machine')
    }
    if ($envPath -and (Test-Path $envPath)) { return $envPath }

    $bases = @($env:ProgramFiles, ${env:ProgramFiles(x86)}) | Where-Object { $_ }
    $direct = foreach ($b in $bases) {
        Join-Path $b 'ODA\ODAFileConverter\ODAFileConverter.exe'
        Join-Path $b 'ODA\ODAFile Converter\ODAFileConverter.exe'
    }
    foreach ($p in $direct) { if (Test-Path $p) { return $p } }

    foreach ($b in $bases) {
        $odaRoot = Join-Path $b 'ODA'
        if (-not (Test-Path $odaRoot)) { continue }
        Get-ChildItem $odaRoot -Directory -ErrorAction SilentlyContinue |
            Where-Object { $_.Name -ilike '*odafileconverter*' } |
            ForEach-Object {
                $exe = Join-Path $_.FullName 'ODAFileConverter.exe'
                if (Test-Path $exe) { return $exe }
            }
    }
    return $null
}

$odaExe = Find-OdaExe

if ($odaExe) {
    Write-OK "Найден: $odaExe"
    # Прописываем переменную окружения для текущего пользователя (если ещё не задана)
    $curEnv = [Environment]::GetEnvironmentVariable('ODA_FILE_CONVERTER_EXE', 'User')
    if ($curEnv -ne $odaExe) {
        [Environment]::SetEnvironmentVariable('ODA_FILE_CONVERTER_EXE', $odaExe, 'User')
        $env:ODA_FILE_CONVERTER_EXE = $odaExe
        Write-OK "ODA_FILE_CONVERTER_EXE прописан в переменные окружения пользователя"
    } else {
        Write-OK "ODA_FILE_CONVERTER_EXE уже задан"
    }
} else {
    # Проверяем — может инсталлятор лежит рядом в папке дистрибутива
    $bundledInstaller = Get-ChildItem $ScriptRoot -Filter 'ODAFileConverter*.exe' |
        Select-Object -First 1

    if ($bundledInstaller) {
        Write-Info "Найден инсталлятор в папке дистрибутива: $($bundledInstaller.Name)"
        Write-Info "Устанавливаем ODA File Converter..."
        try {
            $proc = Start-Process -FilePath $bundledInstaller.FullName `
                -ArgumentList '/S' -Wait -PassThru
            if ($proc.ExitCode -eq 0) {
                $odaExe = Find-OdaExe
                if ($odaExe) {
                    [Environment]::SetEnvironmentVariable('ODA_FILE_CONVERTER_EXE', $odaExe, 'User')
                    $env:ODA_FILE_CONVERTER_EXE = $odaExe
                    Write-OK "ODA File Converter установлен: $odaExe"
                } else {
                    Write-Warn "ODA установлен, но exe не найден — задайте ODA_FILE_CONVERTER_EXE вручную"
                }
            } else {
                Write-Warn "Инсталлятор завершился с кодом $($proc.ExitCode)"
            }
        } catch {
            Write-Warn "Не удалось запустить инсталлятор: $_"
        }
    } else {
        Write-Warn "ODA File Converter не найден."
        Write-Info ""
        Write-Info "  Без него программа не сможет открывать файлы .dwg"
        Write-Info "  (файлы .dxf работают без ODA)"
        Write-Info ""
        Write-Info "  Скачайте бесплатно (требуется регистрация):"
        Write-Info "  https://www.opendesign.com/guestfiles/oda_file_converter"
        Write-Info ""
        Write-Info "  После установки ODA запустите setup.bat повторно,"
        Write-Info "  чтобы переменная ODA_FILE_CONVERTER_EXE была задана автоматически."
        Write-Info ""
        $open = Read-Host "  Открыть страницу загрузки в браузере? (y/n)"
        if ($open -eq 'y' -or $open -eq 'д') {
            Start-Process "https://www.opendesign.com/guestfiles/oda_file_converter"
        }
    }
}

# ── 6. Ярлык на рабочем столе ─────────────────────────────────────────────────
Write-Step "Создание ярлыка на рабочем столе..."

try {
    $wsh = New-Object -ComObject WScript.Shell
    $sc  = $wsh.CreateShortcut($DesktopShortcut)
    $sc.TargetPath       = $ExePath
    $sc.WorkingDirectory = $ScriptRoot
    $sc.Description      = "TechnoPan — генератор спецификации панелей"
    $sc.Save()
    Write-OK "Ярлык создан: $DesktopShortcut"
} catch {
    Write-Warn "Не удалось создать ярлык на рабочем столе: $_"
    Write-Info "Создайте ярлык вручную на $ExePath"
}

# ── 7. Ярлык в меню «Пуск» ────────────────────────────────────────────────────
Write-Step "Добавление в меню Пуск..."

try {
    if (-not (Test-Path $StartMenuDir)) {
        New-Item -ItemType Directory -Path $StartMenuDir | Out-Null
    }
    $wsh = New-Object -ComObject WScript.Shell
    $sc  = $wsh.CreateShortcut($StartMenuLink)
    $sc.TargetPath       = $ExePath
    $sc.WorkingDirectory = $ScriptRoot
    $sc.Description      = "TechnoPan — генератор спецификации панелей"
    $sc.Save()
    Write-OK "Ярлык в Пуске: $StartMenuLink"
} catch {
    Write-Warn "Не удалось создать ярлык в меню Пуск: $_"
}

# ── Итог ─────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "============================================================" -ForegroundColor Green
Write-Host "  Установка завершена!" -ForegroundColor Green
Write-Host "  Запустите $AppName с рабочего стола или через меню Пуск." -ForegroundColor Green
Write-Host "============================================================" -ForegroundColor Green
Write-Host ""
