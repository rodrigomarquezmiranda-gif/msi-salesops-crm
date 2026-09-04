@echo off
cd /d "%~dp0"
echo ============================================
echo   MSI SalesOps CRM - Sincronizando...
echo ============================================
git pull
echo.
echo Listo! Abriendo CRM...
start "" "index.html"
