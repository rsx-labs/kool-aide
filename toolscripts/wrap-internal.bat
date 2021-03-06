@echo off

REM  make-dir.bat
REM  usage: create a package for kool-aide using --onedir params
REM  note: the generated package will be a folder with multiple files

echo kool-aide builder  --onedir
set APP=%1
set VER=%2
set BUILD_FOLDER=%APP%_%VER%_internal

echo building version = %APP%  %VER%

cd ..
if not exist ".\build" (
    echo creating build folder ...
    mkdir "build"
)

cd build

if exist %BUILD_FOLDER% (
    rmdir %BUILD_FOLDER% /Q
)
echo creating app build folder ...
mkdir %BUILD_FOLDER%

echo building %APP% %VER% ....
rd ..\src\kool_aide\cli\dist\kool-aide /s /Q
rd ..\src\kool_aide\cli\build\kool-aide /s /Q
del ..\src\kool_aide\cli\dist\*.* /Q
del ..\src\kool_aide\cli\build\*.* /Q
cd ../src/kool_aide/cli

if not exist ..\..\..\ka-env (
    echo the virtual environment folder does not exist
    Goto ERR
)

pyinstaller --onedir --name %APP% main.py --paths ..\..\..\ka-env\Lib\site-packages --paths ..\..\kool-aide --icon=..\assets\images\kool-aide.ico

cd ../../../build/%BUILD_FOLDER%
mkdir batch
mkdir dbcomp
xcopy ..\..\src\kool_aide\cli\dist\kool-aide /e
copy ..\..\src\kool-aide-settings.prod.json kool-aide-settings.json 
xcopy ..\..\toolscripts\Batch  .\batch /e
xcopy ..\..\dbcomp  .\dbcomp /e
copy ..\..\docs\readme.rtf 
copy ..\..\toolscripts\run-sql.cmd .\dbcomp

cd ..\..
Goto DONE

:ERR
echo ERROR! error building

:DONE
echo build --onedir done