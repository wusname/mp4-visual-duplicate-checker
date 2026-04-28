@echo off
chcp 65001 >nul
setlocal
title MP4 Duplicate Checker v7
set "SCRIPT=%~dp0mp4_duplicate_checker_v7.ps1"
if not exist "%SCRIPT%" (
  echo Cannot find:
  echo %SCRIPT%
  echo Keep BAT and PS1 in the same folder.
  pause
  exit /b 1
)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT%" %*
echo.
pause
exit /b 0
