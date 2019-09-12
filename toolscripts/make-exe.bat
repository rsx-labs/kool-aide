@echo off
echo kool-aide builder
set APP=%1
set VER=%2

echo building version = %VER%

cd ../build

echo building %APP% %VER% ....
del ..\src\kool_aide\dist\*.* /Q
cd ../src/kool_aide/cli

echo %cd%
pyinstaller --onefile --name %APP% main.py --paths ..\..\..\kool-aide-venv\Lib\site-packages --paths ..\..\kool-aide

echo done