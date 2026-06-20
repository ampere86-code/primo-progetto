@echo off
REM Costruisce l'eseguibile Cronocantieri.exe (versione installabile senza Python).
REM Richiede Python e le dipendenze installate (vedi requirements.txt + pyinstaller).
REM Avvio: doppio clic su questo file, oppure da terminale: costruisci_exe.bat

echo Installazione/aggiornamento dipendenze...
python -m pip install -r requirements.txt pyinstaller

echo.
echo Costruzione eseguibile in corso (puo' richiedere qualche minuto)...
python -m PyInstaller --onefile --name Cronocantieri ^
    --collect-all pdfplumber ^
    --collect-all pdfminer ^
    --collect-all openpyxl ^
    main.py

echo.
echo Fatto. L'eseguibile e' in:  dist\Cronocantieri.exe
pause
