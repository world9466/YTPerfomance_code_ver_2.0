REM 有錯誤就寫入error.log
python combine_table.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python combine_table_2.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python KPI_common.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python KPI_common_2.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python KPI_politics.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python KPI_politics_2.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python YTBP_common.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python YTBP_politics.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python audience.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python bangumi.py 2>error.log
if %errorlevel% neq 0 pause exit

REM 沒發生錯誤刪除error.log
if %errorlevel% equ 0 (del /q error.log)
timeout /t 5
