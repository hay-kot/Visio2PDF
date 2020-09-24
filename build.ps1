./venv/Scripts/Activate.ps1

$Version = "v0.6.0"
$ProjName = "Visio2PDF"
$Name = "$ProjName $Version"
$BuildDir = "./dist"
$Build = "$BuildDir/$Name"
$Archive = "$Build.7z"


pyinstaller --noconfirm --windowed `
    --name $Name `
    --add-data="./$ProjName/web/;./web" `
    --add-data="./$ProjName/settings;./settings" `
    --add-binary="./$ProjName/OfficeToPDF.exe;." `
    --icon="./$ProjName/web/favicon.ico" `
    ./$ProjName/$ProjName.py

$7zipPath = "$env:ProgramFiles\7-Zip\7z.exe"

if (-not (Test-Path -Path $7zipPath -PathType Leaf)) {
    throw "7 zip file '$7zipPath' not found"
}

Set-Alias 7zip $7zipPath


Write-Host "Compressing to Archive..."
7zip a $Archive $Build
# Compress-Archive -Path $ArchivePath -DestinationPath "$ArchivePath.zip" -Force
Write-Host "Archive Complete"

