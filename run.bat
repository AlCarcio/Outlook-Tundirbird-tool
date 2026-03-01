@echo off
echo Installing dependencies...
pip install -r requirements.txt --quiet
echo.
echo Starting PST to MBOX Converter...
python main.py
pause
