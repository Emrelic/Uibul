# UIBUL Setup Creator
# Bu script publish klasorunu zipleyip, self-extracting setup olusturur

$ErrorActionPreference = "Stop"
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$projectDir = Join-Path $scriptDir "..\UIElementInspector"
$publishDir = Join-Path $projectDir "publish"
$outputDir = Join-Path $scriptDir "Output"
$desktopPath = [Environment]::GetFolderPath("Desktop")

Write-Host "=== UIBUL Setup Creator ===" -ForegroundColor Cyan

# Step 1: Check publish exists
if (-not (Test-Path $publishDir)) {
    Write-Host "Publish klasoru bulunamadi. Once 'dotnet publish' calistirin." -ForegroundColor Red
    exit 1
}

Write-Host "[1/4] Publish klasoru kontrol edildi: $publishDir" -ForegroundColor Green

# Step 2: Create output directory
if (-not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# Step 3: Create ZIP from publish
$zipPath = Join-Path $outputDir "UIBUL_v3_Files.zip"
if (Test-Path $zipPath) { Remove-Item $zipPath -Force }

Write-Host "[2/4] ZIP dosyasi olusturuluyor..." -ForegroundColor Yellow
Compress-Archive -Path "$publishDir\*" -DestinationPath $zipPath -CompressionLevel Optimal
$zipSize = [math]::Round((Get-Item $zipPath).Length / 1MB, 1)
Write-Host "  -> ZIP olusturuldu: $zipSize MB" -ForegroundColor Green

# Step 4: Create self-extracting setup script
$setupScript = Join-Path $outputDir "UIBUL_v3_Setup.ps1"

$setupContent = @'
# UIBUL v3 - Self Extracting Setup
# Bu script UIBUL uygulamasini kurar

param(
    [string]$InstallDir = "$env:LOCALAPPDATA\UIBUL"
)

$ErrorActionPreference = "Stop"

function Show-Message($msg, $color = "White") { Write-Host $msg -ForegroundColor $color }

Show-Message "=============================================" "Cyan"
Show-Message "   UIBUL - UI Element Inspector v3.0 Setup" "Cyan"
Show-Message "=============================================" "Cyan"
Show-Message ""

# Ask for install location
$defaultDir = "$env:LOCALAPPDATA\UIBUL"
Show-Message "Kurulum dizini: $defaultDir" "Yellow"
$customDir = Read-Host "Farkli bir dizin girmek icin yazin (Enter = varsayilan)"
if ($customDir) { $InstallDir = $customDir }

Show-Message ""
Show-Message "[1/5] Kurulum dizini hazirlaniyor..." "Yellow"

# Create install directory
if (-not (Test-Path $InstallDir)) {
    New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null
}

# Find the zip file (same directory as this script)
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$zipFile = Join-Path $scriptDir "UIBUL_v3_Files.zip"

if (-not (Test-Path $zipFile)) {
    Show-Message "HATA: UIBUL_v3_Files.zip bulunamadi!" "Red"
    Show-Message "Bu dosya setup script ile ayni klasorde olmali." "Red"
    Read-Host "Devam etmek icin Enter'a basin"
    exit 1
}

Show-Message "[2/5] Dosyalar cikartiliyor..." "Yellow"
Expand-Archive -Path $zipFile -DestinationPath $InstallDir -Force
Show-Message "  -> Dosyalar cikartildi" "Green"

# Install Playwright browsers
Show-Message "[3/5] Playwright tarayicilari yukleniyor (bu biraz zaman alabilir)..." "Yellow"
$playwrightPs1 = Join-Path $InstallDir "playwright.ps1"
if (Test-Path $playwrightPs1) {
    try {
        & pwsh $playwrightPs1 install chromium 2>$null
        Show-Message "  -> Playwright Chromium yuklendi" "Green"
    } catch {
        try {
            & powershell -ExecutionPolicy Bypass -File $playwrightPs1 install chromium 2>$null
            Show-Message "  -> Playwright Chromium yuklendi" "Green"
        } catch {
            Show-Message "  -> Playwright yuklenemedi (program yine de calisir, Playwright ozelligi haric)" "DarkYellow"
        }
    }
} else {
    Show-Message "  -> Playwright script bulunamadi, atlaniyor" "DarkYellow"
}

# Create desktop shortcut
Show-Message "[4/5] Masaustu kisayolu olusturuluyor..." "Yellow"
$exePath = Join-Path $InstallDir "UIElementInspector.exe"
$desktopPath = [Environment]::GetFolderPath("Desktop")
$shortcutPath = Join-Path $desktopPath "UIBUL - UI Element Inspector.lnk"
$iconPath = Join-Path $InstallDir "Resources\app.ico"

$WshShell = New-Object -ComObject WScript.Shell
$shortcut = $WshShell.CreateShortcut($shortcutPath)
$shortcut.TargetPath = $exePath
$shortcut.WorkingDirectory = $InstallDir
$shortcut.Description = "UIBUL - Universal UI Element Inspector"
if (Test-Path $iconPath) {
    $shortcut.IconLocation = "$iconPath,0"
}
$shortcut.Save()
Show-Message "  -> Masaustu kisayolu olusturuldu" "Green"

# Create Start Menu shortcut
Show-Message "[5/5] Baslat menusu kisayolu olusturuluyor..." "Yellow"
$startMenuDir = Join-Path $env:APPDATA "Microsoft\Windows\Start Menu\Programs\UIBUL"
if (-not (Test-Path $startMenuDir)) {
    New-Item -ItemType Directory -Path $startMenuDir -Force | Out-Null
}
$startShortcutPath = Join-Path $startMenuDir "UIBUL - UI Element Inspector.lnk"
$startShortcut = $WshShell.CreateShortcut($startShortcutPath)
$startShortcut.TargetPath = $exePath
$startShortcut.WorkingDirectory = $InstallDir
$startShortcut.Description = "UIBUL - Universal UI Element Inspector"
if (Test-Path $iconPath) {
    $startShortcut.IconLocation = "$iconPath,0"
}
$startShortcut.Save()
Show-Message "  -> Baslat menusu kisayolu olusturuldu" "Green"

Show-Message ""
Show-Message "=============================================" "Green"
Show-Message "   KURULUM TAMAMLANDI!" "Green"
Show-Message "=============================================" "Green"
Show-Message ""
Show-Message "Kurulum dizini: $InstallDir" "White"
Show-Message "Masaustu kisayolu olusturuldu" "White"
Show-Message ""

$launch = Read-Host "Uygulamayi simdi baslatmak ister misiniz? (E/H)"
if ($launch -eq "E" -or $launch -eq "e" -or $launch -eq "") {
    Start-Process $exePath
}
'@

Set-Content -Path $setupScript -Value $setupContent -Encoding UTF8
Write-Host "[3/4] Setup script olusturuldu" -ForegroundColor Green

# Step 5: Create a simple BAT launcher (double-click friendly)
$batPath = Join-Path $outputDir "UIBUL_v3_Setup.bat"
$batContent = @"
@echo off
title UIBUL v3 Setup
echo ============================================
echo    UIBUL - UI Element Inspector v3.0 Setup
echo ============================================
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0UIBUL_v3_Setup.ps1"
if %errorlevel% neq 0 (
    echo.
    echo Kurulum sirasinda bir hata olustu.
    pause
)
"@
Set-Content -Path $batPath -Value $batContent -Encoding ASCII

Write-Host "[4/4] BAT launcher olusturuldu" -ForegroundColor Green

# Copy to Desktop
$desktopSetupDir = Join-Path $desktopPath "UIBUL_v3_Setup"
if (Test-Path $desktopSetupDir) { Remove-Item $desktopSetupDir -Recurse -Force }
New-Item -ItemType Directory -Path $desktopSetupDir | Out-Null
Copy-Item $zipPath -Destination $desktopSetupDir
Copy-Item $setupScript -Destination $desktopSetupDir
Copy-Item $batPath -Destination $desktopSetupDir

Write-Host ""
Write-Host "=== TAMAMLANDI ===" -ForegroundColor Green
Write-Host "Setup dosyalari masaustune kopyalandi: $desktopSetupDir" -ForegroundColor Cyan
Write-Host "  - UIBUL_v3_Setup.bat   (cift tikla ve kur)" -ForegroundColor White
Write-Host "  - UIBUL_v3_Setup.ps1   (PowerShell script)" -ForegroundColor White
Write-Host "  - UIBUL_v3_Files.zip   (uygulama dosyalari)" -ForegroundColor White
Write-Host ""
Write-Host "ZIP boyutu: $zipSize MB" -ForegroundColor Yellow
