cls

cd /d %~dp0 
python.exe pkiCMOSDetectorTestScript.py




REM cd "C:\Users\scltester\Documents\GitWorkspace\FaxitronCabinet"
REM python.exe "C:\Users\scltester\Documents\GitWorkspace\FaxitronCabinet\pkiCMOSDetectorTestScript.py"

@echo on

REM IF "%~1"=="" GOTO production


REM python.exe "C:\Users\scltester\Documents\GitWorkspace\FaxitronCabinet\pkiCMOSDetectorTestScript.py"

goto complete


production:
set startDate=%date%
set startTime=%time%

set /a sth=%startTime:~0,2%
set /a stm=1%startTime:~3,2% - 100
set /a sts=1%startTime:~6,2% - 100

python.exe "C:\Users\scltester\Documents\GitWorkspace\FaxitronCabinet\pkiCMOSDetectorTestScript.py" >tmp.txt && type tmp.txt && type tmp.txt > C:\TestData\Logs\log_%startDate%_%sth%.%stm%.%sts%.txt

@echo off

del tmp.txt


complete:
REM: Test Complete!