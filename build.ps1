./venv/Scripts/Activate.ps1

$Version = "v0.41"
$ProjName = "Visio2PDF"
$Name = "$ProjName $Version"
$BuildDir = "./dist"
$ArchivePath = "$BuildDir/$Name"


pyinstaller --noconfirm --windowed `
    --name $Name `
    --add-data="./$ProjName/web/;./web" `
    --add-binary="./$ProjName/OfficeToPDF.exe;." `
    --icon="./$ProjName/web/favicon.ico" `
    ./$ProjName/$ProjName.py

Write-Host "Compressing to Archive..."
Compress-Archive -Path $ArchivePath -DestinationPath "$ArchivePath.zip" -Force
Write-Host "Archive Complete"

