echo off

pip install xlwings
pip install PyQt5
pywin32-221.win-amd64-py3.6.exe
if ERRORLEVEL 1 GOTO NOPYTHON

GOTO :EOF
:NOPYTHON
echo Please Install Python with version 3.6 or higher.
echo https://www.python.org/downloads/release/python-361/
PAUSE