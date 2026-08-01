# ===========================================================================
#  UIBUL - surum uretme betigi
#
#  Uygulamayi self-contained yayinlar, belgeleri ekler, ciktiyi kurulum
#  programinin icine gomer ve tek dosyalik UIBUL_Setup.exe uretir.
#
#  Kullanim:
#    powershell -ExecutionPolicy Bypass -File tools\build-release.ps1
#    ... -SadeceInstaller             (uygulama zaten yayinlandiysa; cok hizli)
#    ... -MasaustuneKopyalama:$false
#
#  NOT: Bu dosya bilerek SAF ASCII yazilmistir. Windows PowerShell 5.1,
#  BOM'suz UTF-8 dosyalari sistem kod sayfasiyla okur; Turkce harfler ve
#  uzun tire gibi karakterler ayristirma hatasina yol aciyordu (olculdu).
# ===========================================================================

[CmdletBinding()]
param(
    [switch] $SadeceInstaller,
    [bool]   $MasaustuneKopyalama = $true
)

$ErrorActionPreference = 'Stop'
$baslangic = Get-Date

# -- Yollar ----------------------------------------------------------------
$kok           = Split-Path -Parent $PSScriptRoot
$uygulamaProj  = Join-Path $kok 'UIElementInspector\UIElementInspector\UIElementInspector.csproj'
$installerProj = Join-Path $kok 'Installer\UibulSetup\UibulSetup.csproj'
$yayinDizin    = Join-Path $kok 'build\publish'
$yukDizin      = Join-Path $kok 'Installer\UibulSetup\Payload'
$yukZip        = Join-Path $yukDizin 'uibul-payload.zip'
$installerCik  = Join-Path $kok 'build\installer'
$dagitim       = Join-Path $kok 'dist'
$belgeler      = Join-Path $kok 'Docs'

function Adim($n, $metin) { Write-Host "`n[$n] $metin" -ForegroundColor Cyan }
function Tamam($metin)    { Write-Host "    OK  $metin" -ForegroundColor Green }
function Bilgi($metin)    { Write-Host "    ..  $metin" -ForegroundColor DarkGray }
function Uyari($metin)    { Write-Host "    !!  $metin" -ForegroundColor Yellow }

Write-Host ""
Write-Host "===========================================" -ForegroundColor Cyan
Write-Host "  UIBUL - surum uretme" -ForegroundColor Cyan
Write-Host "===========================================" -ForegroundColor Cyan

# -- 0. On kosullar --------------------------------------------------------
Adim 0 'On kosullar denetleniyor'

if (-not (Get-Command dotnet -ErrorAction SilentlyContinue)) {
    throw '.NET SDK bulunamadi. https://dotnet.microsoft.com/download adresinden kurun.'
}
Tamam ("dotnet SDK " + (dotnet --version))

foreach ($p in @($uygulamaProj, $installerProj)) {
    if (-not (Test-Path $p)) { throw "Proje bulunamadi: $p" }
}

# Calisan UIBUL derlemeyi kilitler.
Get-Process -Name 'UIElementInspector' -ErrorAction SilentlyContinue | ForEach-Object {
    Uyari "Calisan UIBUL kapatiliyor (pid $($_.Id))"
    try { $_.Kill(); $_.WaitForExit(5000) } catch { }
}

# Surum numarasini csproj'dan oku - tek dogruluk kaynagi orasi.
$surum = ([xml](Get-Content $uygulamaProj)).Project.PropertyGroup.Version |
         Where-Object { $_ } | Select-Object -First 1
if (-not $surum) { $surum = '0.0.0' }
Tamam "Surum: $surum"

# -- 1. Uygulamayi yayinla -------------------------------------------------
if (-not $SadeceInstaller) {
    Adim 1 'Uygulama yayinlaniyor (self-contained, win-x64)'
    Bilgi 'Bu adim birkac dakika surer.'

    if (Test-Path $yayinDizin) { Remove-Item $yayinDizin -Recurse -Force }

    dotnet publish $uygulamaProj `
        -c Release `
        -r win-x64 `
        --self-contained true `
        -p:PublishSingleFile=false `
        -p:DebugType=none `
        -p:GenerateDocumentationFile=false `
        -o $yayinDizin `
        --nologo `
        -v quiet

    if ($LASTEXITCODE -ne 0) { throw "dotnet publish basarisiz (kod $LASTEXITCODE)" }

    $exe = Join-Path $yayinDizin 'UIElementInspector.exe'
    if (-not (Test-Path $exe)) { throw "Yayin ciktisinda exe yok: $exe" }

    $mb = [math]::Round(((Get-ChildItem $yayinDizin -Recurse -File |
           Measure-Object Length -Sum).Sum / 1MB), 1)
    Tamam "Yayin tamam - $mb MB"
}
else {
    Adim 1 'Uygulama yayini atlandi (-SadeceInstaller)'
    if (-not (Test-Path (Join-Path $yayinDizin 'UIElementInspector.exe'))) {
        throw 'Onceki yayin ciktisi yok. -SadeceInstaller olmadan calistirin.'
    }
    Tamam 'Mevcut yayin kullanilacak'
}

# -- 2. Belgeleri ekle -----------------------------------------------------
Adim 2 'Belgeler ciktiya kopyalaniyor'

$hedefBelge = Join-Path $yayinDizin 'Docs'
New-Item -ItemType Directory -Path $hedefBelge -Force | Out-Null

if (Test-Path $belgeler) {
    Copy-Item (Join-Path $belgeler '*') -Destination $hedefBelge -Recurse -Force
    $sayi = (Get-ChildItem $hedefBelge -Recurse -File).Count
    Tamam "$sayi belge dosyasi eklendi"
}
else {
    Uyari 'Docs klasoru yok - belgeler atlandi'
}

# Eski kullanim kilavuzu da gitsin. Dosya adinda Turkce harf var; bu betik
# ASCII oldugu icin ada gore degil, uzantiya gore bulunuyor.
Get-ChildItem -Path (Join-Path $kok 'UIElementInspector') -Filter '*.txt' -File `
    -ErrorAction SilentlyContinue |
    ForEach-Object { Copy-Item $_.FullName -Destination $hedefBelge -Force }

# -- 3. Yuku paketle -------------------------------------------------------
Adim 3 'Uygulama paketi (yuk) sikistiriliyor'

New-Item -ItemType Directory -Path $yukDizin -Force | Out-Null
if (Test-Path $yukZip) { Remove-Item $yukZip -Force }

# .NET'in kendi ZipFile'i Compress-Archive'dan hizlidir ve buyuk dosyalarda
# takilmaz. Optimal sikistirma setup boyutunu belirgin dusuruyor.
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory(
    $yayinDizin, $yukZip,
    [System.IO.Compression.CompressionLevel]::Optimal, $false)

$zipMb = [math]::Round((Get-Item $yukZip).Length / 1MB, 1)
Tamam "Yuk hazir - $zipMb MB"

# -- 4. Installer'i derle --------------------------------------------------
Adim 4 'Kurulum programi derleniyor'

if (Test-Path $installerCik) { Remove-Item $installerCik -Recurse -Force }

dotnet publish $installerProj `
    -c Release `
    -o $installerCik `
    --nologo `
    -v quiet

if ($LASTEXITCODE -ne 0) { throw "Installer derlemesi basarisiz (kod $LASTEXITCODE)" }

$setupExe = Join-Path $installerCik 'UIBUL_Setup.exe'
if (-not (Test-Path $setupExe)) { throw "Setup exe uretilemedi: $setupExe" }

$setupMb = [math]::Round((Get-Item $setupExe).Length / 1MB, 1)
Tamam "UIBUL_Setup.exe - $setupMb MB"

# Yukun gercekten gomuldugunu dogrula: setup, zip'ten belirgin buyuk olmali.
if ((Get-Item $setupExe).Length -lt (Get-Item $yukZip).Length) {
    throw 'Setup dosyasi yukten kucuk - paket gomulmemis olabilir. Payload yolunu denetleyin.'
}

# -- 5. Dagitim klasoru ----------------------------------------------------
Adim 5 'Dagitim klasoru hazirlaniyor'

New-Item -ItemType Directory -Path $dagitim -Force | Out-Null
Get-ChildItem $dagitim -Filter 'UIBUL_Setup*.exe' -ErrorAction SilentlyContinue |
    Remove-Item -Force -ErrorAction SilentlyContinue

$dagitimExe = Join-Path $dagitim 'UIBUL_Setup.exe'
$surumluExe = Join-Path $dagitim "UIBUL_Setup_v$surum.exe"

Copy-Item $setupExe -Destination $dagitimExe -Force
Copy-Item $setupExe -Destination $surumluExe -Force
Tamam 'dist\UIBUL_Setup.exe'
Tamam "dist\UIBUL_Setup_v$surum.exe"

# SHA256 - arkadasiniz dosyanin bozulmadigini dogrulayabilsin.
$ozet = (Get-FileHash $dagitimExe -Algorithm SHA256).Hash
"$ozet  UIBUL_Setup_v$surum.exe" |
    Set-Content (Join-Path $dagitim 'SHA256.txt') -Encoding utf8
Tamam ("SHA256: " + $ozet.Substring(0, 16) + "...")

# -- 6. Masaustune kopyala -------------------------------------------------
if ($MasaustuneKopyalama) {
    Adim 6 'Masaustune kopyalaniyor'
    $masaustu = [Environment]::GetFolderPath('Desktop')
    $masaustuExe = Join-Path $masaustu 'UIBUL_Setup.exe'

    try {
        Copy-Item $dagitimExe -Destination $masaustuExe -Force
        Tamam $masaustuExe
    }
    catch {
        Uyari ("Masaustune kopyalanamadi: " + $_.Exception.Message)
    }
}

# -- Ozet ------------------------------------------------------------------
$sure = [math]::Round(((Get-Date) - $baslangic).TotalSeconds)

Write-Host ""
Write-Host "===========================================" -ForegroundColor Green
Write-Host "  TAMAMLANDI  ($sure sn)" -ForegroundColor Green
Write-Host "===========================================" -ForegroundColor Green
Write-Host ""
Write-Host "  Surum      : $surum"
Write-Host "  Setup      : $dagitimExe  ($setupMb MB)"
if ($MasaustuneKopyalama) { Write-Host "  Masaustu   : UIBUL_Setup.exe" }
Write-Host ""
Write-Host "  Siradaki adim - GitHub'a yayinlayin:" -ForegroundColor Yellow
Write-Host "    gh release create v$surum dist/UIBUL_Setup.exe --title UIBUL-v$surum --notes ..." -ForegroundColor Gray
Write-Host ""
Write-Host "  Ayrinti: Docs\SURUM-CIKARMA.md" -ForegroundColor DarkGray
Write-Host ""
