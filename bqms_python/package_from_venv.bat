@echo off
chcp 65001 >nul

echo venv offline packaging...

call venv\Scripts\activate.bat
echo venv activated

python create_offline_package.py

echo Done!
pause