@echo off
start "" /min powershell -WindowStyle Hidden -NoProfile -ExecutionPolicy Bypass -File "%~dp0ImageConverter-PS-GUI.ps1"
exit
