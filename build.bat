@echo off
setlocal

python -m venv .venv
call .venv\Scripts\activate

pip install --upgrade pip
pip install -r requirements.txt
pip install pyinstaller

pyinstaller -w --name SerialCmdTester --clean main.py

echo.
echo Build done. EXE is in: dist\SerialCmdTester\SerialCmdTester.exe
pause
