@echo off
cd /d "%~dp0"
if exist "%~dp0ledger.exe" (
  start "" "%~dp0ledger.exe"
) else (
  start "" "%~dp0index.html"
)
