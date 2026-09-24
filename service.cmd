@echo off
rem Starts relaycheck in the background (no window) and opens the dashboard at http://localhost:8765.
rem Runs a check now and every 30 minutes. Stop it with the Stop Service button at the top of the dashboard.
start "" powershell -NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -File "%~dp0relaycheck.ps1" -Serve -Loop 30
