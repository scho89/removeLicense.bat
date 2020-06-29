@echo off
echo Searching Office installation path....

cd "c:\Program Files (x86)"
dir /s ospp.vbs | findstr /r "Office" >> %temp%\__osppPath.txt

cd "c:\Program Files"
dir /s ospp.vbs | findstr /r "Office" >> %temp%\__osppPath.txt

set /p oPath=<%temp%\__osppPath.txt
set oPath=%oPath: 디렉터리=%
set oPath=%oPath: Directory=%
cd %oPath%

echo Removing KMS host information....

cscript ospp.vbs /remhst

echo Getting Office product keys....
cscript ospp.vbs /dstatusall | findstr Last > %temp%\__osppResult.txt

set txtfile=%temp%\__osppResult.txt
set newfile=%temp%\__osppKey.txt
set search=Last 5 characters of installed product key: 
set replace=

setlocal enabledelayedexpansion

for /f "tokens=*" %%a in (%temp%\__osppResult.txt) do (
   set newline=%%a
   set newline=!newline:%search%=%replace%!
   echo !newline! >> %newfile%
)

for /f "tokens=*" %%b in (%temp%\__osppKey.txt) do (
    echo Removing Office product key : %%b ......
    cscript ospp.vbs /unpkey:%%b
)


del %temp%\__ospp*
