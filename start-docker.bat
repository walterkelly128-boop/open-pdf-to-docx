@echo off
setlocal
cd /d "%~dp0"
echo Starting Open PDF to DOCX with Docker Desktop...
docker compose up --build
