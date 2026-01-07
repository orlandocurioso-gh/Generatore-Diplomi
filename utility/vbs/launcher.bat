@echo off
SETLOCAL

REM ----------------------------------------------------
REM 1. AVVIA IL SERVER FLASK
REM ----------------------------------------------------

SET FLASK_APP_PATH="C:\xampp\htdocs\generatorediplomi6.9\app.py"
start /min cmd /c python %FLASK_APP_PATH%

REM ----------------------------------------------------
REM 2. ATTENDI L'AVVIO DEL SERVER
REM ----------------------------------------------------

timeout /t 5 /nobreak >nul

REM ----------------------------------------------------
REM 3. APRI LA PAGINA NEL BROWSER A TUTTO SCHERMO
REM ----------------------------------------------------

SET CHROME_PATH="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
SET TARGET_URL="http://127.0.0.1:5000/"

REM *** MODIFICA QUI: Aggiunto l'argomento --start-fullscreen ***
start "" %CHROME_PATH% %TARGET_URL% --new-window

REM --start-fullscreen

REM Se preferisci la modalità KIOSK completa (nasconde anche la barra di Windows):
REM start "" %CHROME_PATH% %TARGET_URL% --kiosk

REM start "" %CHROME_PATH% --app=%TARGET_URL% --kiosk

exit