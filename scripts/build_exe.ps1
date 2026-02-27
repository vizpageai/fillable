param(
    [string]$Python = "python"
)

& $Python -m pip install --upgrade pip
& $Python -m pip install -r requirements.txt
& $Python -m pip install pyinstaller

& $Python -m PyInstaller `
    --noconfirm `
    --onefile `
    --windowed `
    --name Fillable `
    --collect-all docx `
    --collect-all pptx `
    run_fillable.py

Write-Host "Built dist\\Fillable.exe"
