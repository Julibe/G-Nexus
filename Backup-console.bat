@echo off
title Google Drive Backup (menu)
mode con: cols=200 lines=50
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0backup-drive.ps1"
pause
