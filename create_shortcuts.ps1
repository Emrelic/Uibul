# Masaustu kisayolu olusturur.
#
# Hedef exe SABIT YAZILMAZ — hedef cerceve klasoru (net8.0-windows,
# net10.0-windows ...) surum yukseltmesinde degisiyor ve kisayol sessizce
# bozuluyor. Yasandi: proje net10.0'a gecince kisayol net8.0'i gostermeye
# devam etti, tiklayinca hicbir sey olmadi. Bu yuzden en yeni derlenmis
# exe her calistirmada aranir.

$proje = "C:\Users\emrem\OneDrive\Belgeler\Projeler\Uibul\UIElementInspector\UIElementInspector"
$iconPath = Join-Path $proje "Resources\app.ico"

$exe = Get-ChildItem -Path (Join-Path $proje "bin") -Recurse -Filter "UIElementInspector.exe" -ErrorAction SilentlyContinue |
       Sort-Object LastWriteTime -Descending |
       Select-Object -First 1

if ($null -eq $exe) {
    Write-Host "HATA: derlenmis exe bulunamadi. Once 'dotnet build' calistirin." -ForegroundColor Red
    exit 1
}

$exePath = $exe.FullName
Write-Host "Hedef exe: $exePath"

$desktopPath = [System.Environment]::GetFolderPath('Desktop')
$shell = New-Object -ComObject WScript.Shell

$shortcut = $shell.CreateShortcut("$desktopPath\UI Element Inspector.lnk")
$shortcut.TargetPath = $exePath
$shortcut.IconLocation = $iconPath
$shortcut.WorkingDirectory = Split-Path $exePath
$shortcut.Description = "UI Element Inspector - F11: Tarih Atlasi karesi"
$shortcut.Save()
Write-Host "Masaustu kisayolu olusturuldu: $desktopPath\UI Element Inspector.lnk"

# Eski/bozuk adla duran kisayol varsa temizle
$eski = "$desktopPath\UIElementInspector.lnk"
if (Test-Path $eski) {
    Remove-Item $eski -Force
    Write-Host "Eski (bozuk) kisayol silindi: $eski"
}
