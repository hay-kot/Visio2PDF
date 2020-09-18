./venv/Scripts/Activate.ps1

$Version = "v0.32"
$Name = "Visio2PDF $Version"
$BuildDir = "./dist"
$ArchivePath = "$BuildDir/$Name"


pyinstaller --noconfirm --windowed `
    --name $Name `
    --add-data="./Visio2PDF/web/;./web" `
    --add-binary="./Visio2PDF/OfficeToPDF.exe;." `
    --icon="./Visio2PDF/web/favicon.ico" `
    ./Visio2PDF/Visio2PDF.py

Write-Host "Compressing to Archive..."
Compress-Archive -Path $ArchivePath -DestinationPath "$ArchivePath.zip" -Force
Write-Host "Archive Complete"

