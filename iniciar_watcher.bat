@echo off
:: ============================================================
:: iniciar_watcher.bat
:: Inicia el watcher en segundo plano al abrir esta ventana.
:: Puedes poner un acceso directo a este .bat en el Inicio de
:: Windows para que arranque automáticamente al encender el PC.
:: ============================================================

title BVS Dashboard Watcher

echo.
echo  ============================================
echo   BVS Dashboard - Watcher de Excel
echo  ============================================
echo.
echo  Monitoreando cambios en el Excel...
echo  Cuando guardes el archivo, el dashboard
echo  se actualizara automaticamente en GitHub.
echo.
echo  No cierres esta ventana mientras trabajas.
echo  Presiona Ctrl+C para detener.
echo.

:: Cambia esta ruta a donde tengas Python instalado
:: (normalmente no necesitas cambiarlo si Python está en el PATH)
python "%~dp0watch_and_push.py"

pause
