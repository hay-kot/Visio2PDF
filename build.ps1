./venv/Scripts/Activate.ps1

pyinstaller --noconfirm --windowed `
    --add-data="./Visio2PDF/web/;./web" `
    --add-binary="./Visio2PDF/OfficeToPDF.exe;." `
    --icon="./Visio2PDF/web/favicon.ico" `
    ./Visio2PDF/Visio2PDF.py