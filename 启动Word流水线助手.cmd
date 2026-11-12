@echo off
setlocal
title Word Pipeline Assistant
set DSH_WORD_AUTO_PANEL=1
if exist "C:\Program Files\LibreOffice\program\soffice.exe" set "DSH_WORD_SOFFICE=C:\Program Files\LibreOffice\program\soffice.exe"
if exist "C:\Program Files (x86)\LibreOffice\program\soffice.exe" set "DSH_WORD_SOFFICE=C:\Program Files (x86)\LibreOffice\program\soffice.exe"
start "" /min cmd /c "D:\DSH\resources\runtime\cli\bin\dsh.cmd --profile word --no-open"
start "" "D:\desktop\project\DSH-word\plugin\loading.html"