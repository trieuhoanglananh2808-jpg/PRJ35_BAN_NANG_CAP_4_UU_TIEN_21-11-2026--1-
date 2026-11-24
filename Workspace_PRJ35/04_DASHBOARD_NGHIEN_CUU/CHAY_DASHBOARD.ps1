# Chạy trong PowerShell; không cần kích hoạt môi trường ảo.
$prjPython = Join-Path $env:LOCALAPPDATA 'PRJ35_Streamlit_venv/Scripts/python.exe'
if (-not (Test-Path -LiteralPath $prjPython)) {
    Write-Host 'Chua co moi truong. Xem HUONG_DAN.txt de cai dat Python va requirements.txt.'
    exit 1
}
Set-Location -LiteralPath $PSScriptRoot
& $prjPython -m streamlit run app.py --server.address 127.0.0.1
