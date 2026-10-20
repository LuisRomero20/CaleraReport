@echo off
chcp 65001 > nul
title CALERA - Renombrar Documentos con MIRO
echo ===================================================
echo     CALERA - RENOMBRAR DOCUMENTOS CON MIRO
echo ===================================================
echo.
echo Iniciando herramienta de renombrado con MIRO...
python scripts\renombrar_documentos_miro.py
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] No se pudo iniciar el script de renombrado.
    pause
)
