@echo off
chcp 65001 >nul
cd /d "%~dp0"
"C:\Users\34862\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" "%~dp0weighing_tool_remake.py"
pause
