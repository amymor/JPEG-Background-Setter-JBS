@echo off & cd /d "%~dp0"
fsutil dirty query %systemdrive% >nul && goto:GA || Modules\nsudo -U:E -P:E -UseCurrentConsole -ShowWindowMode:show "%~0" %* && exit /b
:GA
echo.
:Set-FolderPath
:: echo [33m Slideshow FolderPath [0m
:: Check if %1 is provided, if not, read from FolderPath.ini
IF "%~1"=="" (set /p FolderPath=<FolderPath.ini) Else (SET "FolderPath=%~1")

IF "%FolderPath%"=="" (echo. Pls select your desired folder & for /f "tokens=* delims=" %%p in ('cscript //nologo Modules\folderpicker.vbs') do set "FolderPath=%%p")

:: test FolderPath variable
echo %FolderPath%

:Creat-FolderPath.ini
:: store folder path for desktop context menu
echo %FolderPath%>FolderPath.ini

:list
:: make previous used list.ini empty
type NUL > list.ini

echo [33m Creating a list of images file (this may take some time depending on the number of images and the CPU's power) [0m
ListFiles.exe "%FolderPath%">> list.ini

:LineNumber
SET /P "LineNumber= Enter the line number to use as a variable. If you want to set first image as background, then enter 1:"

:: create ini to store number of currnet background
IF NOT EXIST "CurrentBG*" (
  type NUL > CurrentBG1
)
 :: rename previous CurrentBG* to specified number:
REN "CurrentBG*" "CurrentBG%LineNumber%"
 :: find the .ini file and set it as a variable:
FOR /R "%~dp0" %%F IN (CurrentBG*) DO SET "IniFile=%%~nF"
 :: remove the CurrentBG
SET "LineNumber=%IniFile:CurrentBG=%"

:: echo %LineNumber%

SETLOCAL EnableDelayedExpansion
SET LineCount=0
FOR /F "delims=" %%L IN (list.ini) DO (
  SET /A LineCount+=1
  IF !LineCount! EQU %LineNumber% (
    SET "Variable=%%L"
    GOTO BreakLoop
  )
)
:BreakLoop
:: echo "%Variable%"
IF "%Variable%"=="" (echo Error seems you entered 0 so the variable is empty && echo Press Enter to exit... && pause && exit) Else (goto set-background)

:set-background
echo "%FolderPath%\%Variable%"
"%~dp0JBS.exe" "%FolderPath%\%Variable%"

echo [92m Done! [0m

echo.
pause
exit