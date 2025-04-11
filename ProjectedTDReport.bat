@echo off
REM Change to the directory containing the scripts
python -c "import os; print(os.path.join(os.path.expanduser('~'), 'Desktop', 'ProjectedTD', 'scripts'))" > temp_path.txt
set /p SCRIPT_DIR=<temp_path.txt
del temp_path.txt
cd "%SCRIPT_DIR%"

REM Get current date and time
echo [%date% %time%] Starting script execution... >> log.txt
set "START_TIME=%time%"

REM Run the first Python script and log output to log.txt while showing prints in the terminal
set "TIMER_START=%time%"
echo [%date% %time%] Running queryntr(ntr_tbl).py... >> log.txt
python "queryntr(ntr_tbl).py"
echo [%date% %time%] Finished running queryntr(ntr_tbl).py >> log.txt

REM Calculate elapsed time for queryntr(ntr_tbl).py
set "TIMER_END=%time%"
call :ElapsedTime "%TIMER_START%" "%TIMER_END%"
echo [%date% %time%] Time lapsed for queryntr(ntr_tbl).py: %ELAPSED_TIME% >> log.txt
echo [%date% %time%] 33%% Complete: queryntr(ntr_tbl).py done. Time lapsed: %ELAPSED_TIME%.

REM Run the second Python script and log output to log.txt while showing prints in the terminal
set "TIMER_START=%time%"
echo [%date% %time%] Running copyBRFile.py... >> log.txt
python copyBR_exact.py
echo [%date% %time%] Finished running copyBRFile.py >> log.txt

REM Calculate elapsed time for copyBRFile.py
set "TIMER_END=%time%"
call :ElapsedTime "%TIMER_START%" "%TIMER_END%"
echo [%date% %time%] Time lapsed for copyBRFile.py: %ELAPSED_TIME% >> log.txt
echo [%date% %time%] 66%% Complete: copyBRFile.py done. Time lapsed: %ELAPSED_TIME%.

REM Run the third Python script (rename.py) and log output to log.txt while showing prints in the terminal
set "TIMER_START=%time%"
echo [%date% %time%] Running rename.py... >> log.txt
python rename.py
echo [%date% %time%] Finished running rename.py >> log.txt

REM Calculate elapsed time for rename.py
set "TIMER_END=%time%"
call :ElapsedTime "%TIMER_START%" "%TIMER_END%"
echo [%date% %time%] Time lapsed for rename.py: %ELAPSED_TIME% >> log.txt
echo [%date% %time%] 100%% Complete: rename.py done. Time lapsed: %ELAPSED_TIME%.

REM Calculate total elapsed time
set "END_TIME=%time%"
call :ElapsedTime "%START_TIME%" "%END_TIME%"
echo [%date% %time%] Total execution time: %ELAPSED_TIME% >> log.txt
echo Total execution time: %ELAPSED_TIME%.

REM Pause to keep the console window open
pause

REM Subroutine to calculate elapsed time
:ElapsedTime
setlocal
for /f "tokens=1-4 delims=:.," %%a in ("%~1") do set /a "startTime=((%%a*60+1%%b-100)*60+1%%c-100)*100+1%%d-100"
for /f "tokens=1-4 delims=:.," %%a in ("%~2") do set /a "endTime=((%%a*60+1%%b-100)*60+1%%c-100)*100+1%%d-100"
set /a elapsedTime=endTime-startTime
set /a hours=elapsedTime/360000
set /a minutes=(elapsedTime-hours*360000)/6000
set /a seconds=(elapsedTime-hours*360000-minutes*6000)/100
set /a centiseconds=elapsedTime%%100
endlocal & set "ELAPSED_TIME=%hours%h %minutes%m %seconds%s %centiseconds%cs"
exit /b
