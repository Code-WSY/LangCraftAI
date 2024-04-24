@echo off
echo Please select the function you want to use:
echo [1] Translate
echo [2] PDF to TXT
echo.

:choice
set /P C=ENTER:
if "%C%"=="1" goto :translate
if "%C%"=="2" goto :pdftotxt
echo Please select a valid option
goto choice

:translate
cd /D C:\Users\suyun\OneDrive\Project\LangCraftAI\src
D:\software\anaconda3\envs\API\python.exe GUI.py
goto end

:pdftotxt
cd /D C:\Users\suyun\OneDrive\Project\LangCraftAI\test
D:\software\anaconda3\envs\API\python.exe pdftotxt.py
goto end

:end
pause