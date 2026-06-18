@echo off
echo Generando Motor Dashboard...

wsl bash -lc "cd ~/projects/engine-programme && source .venv/bin/activate && python engine_kpi_report.py"

echo.
echo Reporte generado. Presione una tecla para cerrar.
pause
