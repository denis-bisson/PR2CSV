@echo off
SET THEFILE=C:\Users\explo\OneDrive\FakeDrive\Documents\Lazarus\pr2csv\src\PR2CSV.exe
echo Linking %THEFILE%
C:\lazarus\fpc\3.2.0\bin\x86_64-win64\ld.exe -b pei-x86-64  --gc-sections   --subsystem windows --entry=_WinMainCRTStartup    -o C:\Users\explo\OneDrive\FakeDrive\Documents\Lazarus\pr2csv\src\PR2CSV.exe C:\Users\explo\OneDrive\FakeDrive\Documents\Lazarus\pr2csv\src\link.res
if errorlevel 1 goto linkend
goto end
:asmend
echo An error occurred while assembling %THEFILE%
goto end
:linkend
echo An error occurred while linking %THEFILE%
:end
