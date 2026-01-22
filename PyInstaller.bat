@ECHO OFF
COLOR 0A
SET PYTHONHOME=
SET PYTHONPATH=
SET VENV_PATH=%~dp0
CD /D %VENV_PATH%
CALL %VENV_PATH%.venv\Scripts\activate.bat
python -m PyInstaller -F -n Excel合并工具 --console main.py
PAUSE
EXIT
