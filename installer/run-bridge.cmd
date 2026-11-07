@echo off
setlocal
if "%CYBERA_BRIDGE_API_KEY%"=="" (
  echo CYBERA_BRIDGE_API_KEY is not set for this process.
  echo Follow the installed Bridge Setup Guide before starting the service.
  exit /b 1
)
"%~dp0Cybera.ModernBridge.exe" --urls "http://0.0.0.0:5088"
