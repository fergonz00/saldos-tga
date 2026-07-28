@echo off
REM Vuelca las facturas nuevas de la planilla de Monica a la solapa Compras VW
REM de saldos.titogonzalez.online. Lo dispara la tarea programada "TGA-SyncComprasMonica".
cd /d "%~dp0"
if not exist logs mkdir logs
"C:\Program Files\nodejs\node.exe" sync-compras-monica.js >> "logs\run.log" 2>&1
echo. >> "logs\run.log"
