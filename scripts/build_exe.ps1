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
    --name FillableDOC `
    --version-file packaging\windows_version_info.txt `
    --collect-all docx `
    --collect-all pptx `
    run_fillable.py

Write-Host "Built dist\\FillableDOC.exe"
