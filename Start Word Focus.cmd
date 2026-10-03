@echo off
rem Starts WordFocus.ps1 (same folder) with no console window.
rem -ExecutionPolicy Bypass applies to this one run only; it does not change any system setting.
start "" powershell.exe -NoProfile -STA -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0WordFocus.ps1"
