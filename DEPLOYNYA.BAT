SET D=S:
SET P=NYA
SET F=bin\x86\Release

SET PROJECTS=ABSCS ABSC ABSX AP AR AS ED GL IC PO SA SO TA WH WO TAC
XCOPY ABS\%F%\*.EXE bin /Y /D /C
rem XCOPY ABS\%F%\*.EXE binDebug /Y /D /C
rem XCOPY ABS\%F%\*.PDB binDebug /Y /D /C
XCOPY ABS\%F%\*.ISL bin /Y /D /C
DEL Temp\*.* /q
for %%i in (%PROJECTS%) DO (
DEL %%i\Datasets\*.* /q
XCOPY %%i\%F%\*.DLL bin /Y /D /C
rem XCOPY %%i\%F%\*.DLL binDebug /Y /D /C
rem XCOPY %%i\%F%\*.PDB binDebug /Y /D /C
)
XCOPY bin %D%\%P%\bin /Y /E /C /D >> deploy.log
XCOPY ..\VDI %D%\%P% /Y /E /C /D /exclude:deploy.exclude >> deploy.log
PAUSE