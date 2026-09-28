@echo off
cd /d "%~dp0"
echo Setting things up, please wait...
py -m pip install -r requirements.txt
echo.
echo Starting the app... Leave this window open.
echo In a second, open your browser and go to: localhost:5000
echo.
py app.py
pause
