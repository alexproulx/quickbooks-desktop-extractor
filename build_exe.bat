@echo off
REM Build QBExtract.exe — standalone QuickBooks data extractor
REM No Python needed on the client machine after building
REM
REM Prerequisites (on YOUR dev machine only):
REM   pip install pyinstaller pywin32
REM
REM Output: dist\QBExtract.exe (~8-12 MB single file)

echo Building QBExtract.exe...
pyinstaller --onefile --console --name QBExtract QBExtract.py

if %ERRORLEVEL% EQU 0 (
    echo.
    echo SUCCESS: dist\QBExtract.exe is ready
    echo Send this single file to the client.
    echo They just double-click it with QuickBooks open.
) else (
    echo.
    echo BUILD FAILED — make sure pyinstaller and pywin32 are installed:
    echo   pip install pyinstaller pywin32
)
pause
