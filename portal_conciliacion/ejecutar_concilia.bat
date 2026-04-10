@echo off
cd /d C:\middleware_conta
"C:\Users\Administrator\AppData\Local\Programs\Python\Python313\python.exe" "concilia_sql_postgres.py"
echo.
echo ==========================================
echo Proceso finalizado
echo ==========================================
echo.
echo Presiona cualquier tecla para cerrar o espera 5 segundos...
timeout /t 5