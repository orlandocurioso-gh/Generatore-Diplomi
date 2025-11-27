@echo off
SETLOCAL

REM ----------------------------------------------------
REM 1. AVVIA IL SERVER FLASK
REM ----------------------------------------------------

REM Definisce il percorso dell'applicazione Python
SET FLASK_APP_PATH="C:\xampp\htdocs\generatorediplomi6.7\app.py"

REM Avvia il server Flask in una NUOVA finestra del terminale
REM Questo permette al server di continuare a funzionare in background.
echo Avvio del server Flask...
start "Flask Server" cmd /k python %FLASK_APP_PATH%

REM ----------------------------------------------------
REM 2. ATTENDI L'AVVIO DEL SERVER
REM ----------------------------------------------------

REM Attendi 5 secondi. Regola questo valore se il tuo server impiega più tempo.
echo Attendo 5 secondi per l'avvio del server...
timeout /t 5 /nobreak >nul

REM ----------------------------------------------------
REM 3. APRI LA PAGINA NEL BROWSER
REM ----------------------------------------------------

SET CHROME_PATH="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
SET TARGET_URL="http://127.0.0.1:5000/"

echo Apertura di Chrome...
start "" %CHROME_PATH% %TARGET_URL%

REM ----------------------------------------------------
REM FINE
REM ----------------------------------------------------

ENDLOCAL
REM NOTA: Non usare 'exit' qui, altrimenti chiuderesti subito la finestra principale
REM e non vedresti i messaggi di output. Premi un tasto per chiudere.
pause