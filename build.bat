echo off
rd /s /Q dist\
pyinstaller execl2lua.py -F
copy /y cfg.ini dist\
copy /y alise.txt dist\
pause