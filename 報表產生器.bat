REM anyerror��Ƨ����s�b�N�Ы�
REM if exist anyerror (echo directory CHECK OK) else (mkdir anyerror)

REM �p�G���~�^������0�h�Ȱ�
python combine_table.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python combine_table_2.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python KPI.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python KPI_2.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python YTBP.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python audience.py 2>error.log
if %errorlevel% neq 0 pause exit
timeout /t 2

python bangumi.py 2>error.log
if %errorlevel% neq 0 pause exit

REM �p�G���~�^����0�A�h�R��error.log
if %errorlevel% equ 0 (del /q error.log)
timeout /t 5
