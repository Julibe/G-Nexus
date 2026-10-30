@echo off
title Google Drive Backup
start "" powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0backup-app.ps1"
