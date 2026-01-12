@echo off
title Avvio Gestionale Ordini
color 0A

echo ======================================================
echo      AVVIO DEL GESTIONALE ORDINI IN CORSO...
echo ======================================================
echo.
echo 1. Sto accendendo il server...
cd /d "%~dp0"

:: Avvia Python ridotto a icona
start "GestionaleServer" /min py app.py

echo 2. Attendi qualche secondo che si carichi tutto...
:: Qui aspetta 15 secondi (il >nul nasconde il conto alla rovescia brutto)
timeout /t 15 /nobreak >nul

echo 3. Apro il browser...
start http://127.0.0.1:5000

:: Chiude
exit