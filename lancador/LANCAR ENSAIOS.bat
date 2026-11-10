@echo off
rem Lanca os ensaios exportados pelo app nas fichas e no Resumo (veja config.json)
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Lancar-Ensaios.ps1"
