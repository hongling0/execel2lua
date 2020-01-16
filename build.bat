echo off
cd src
rd /s /Q ..\dist
pyinstaller execl2lua.py -F --distpath=..\dist
copy /y cfg.ini ..\dist
copy /y alise.txt ..\dist
cd ..
pause