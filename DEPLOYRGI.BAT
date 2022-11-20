SET D=S:
SET P=RGI
SET F=bin\x86\Release

SET PROJECTS=ABSCS ABSC ABSX AP AR AS EC ED GL IC PO SA SO TA WB WH WO TAC
XCOPY ABS\%F%\*.EXE bin /Y /D /C
XCOPY ABS\%F%\*.ISL bin /Y /D /C
DEL Temp\*.* /q
for %%i in (%PROJECTS%) DO (
DEL %%i\Datasets\*.* /q
XCOPY %%i\%F%\*.DLL bin /Y /D /C
)
XCOPY bin %D%\%P%\bin /Y /E /C /D >> deploy.log
XCOPY ..\VDI %D%\%P% /Y /E /C /D /exclude:deploy.exclude >> deploy.log
PAUSE