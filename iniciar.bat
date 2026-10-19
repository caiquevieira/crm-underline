@echo off
title CRM
cd /d "%~dp0"
where python >nul 2>nul || (echo Python nao encontrado. Instale em https://www.python.org e marque "Add to PATH". & pause & exit /b 1)
python -c "import openpyxl" 2>nul || python -m pip install openpyxl
python server.py
if errorlevel 1 pause
