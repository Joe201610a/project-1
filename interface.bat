@echo off
setlocal enabledelayedexpansion

set "names=Item-names.txt"
set "prices=Item-prices.txt"
set "reqdata=requested-data.txt"
set "reqquantity=requested-quantity.txt"

set "count=0"

for /f "delims=" %%a in (!names!) do (
    set /a count+=1
    set "name[!count!]=%%a"
)

set "count=0"

for /f "delims=" %%a in (!prices!) do (
    set /a count+=1
    set "price[!count!]=%%a"
)
set "rcount=1"

:loop
cls
echo - choose from below - (press 0 to finish)
echo.

for /l %%x in (1,1,!count!) do (
echo %%x - !name[%%x]! : !price[%%x]! 
)

set /p r=""

if !r! == 0 goto next

echo Quantity:
set /p q=""
set "rdata[!rcount!]=!name[%r%]!"
set "rquantity[!rcount!]=!q!"
set /a rcount+=1
goto loop

:next
set /a rcount-=1
cls
echo Requested data - name : quantity
echo.
for /l %%x in (1,1,!rcount!) do (
set "temp=!rdata[%%x]!"
echo %%x - !rdata[%%x]! : !rquantity[%%x]!
)
echo.
echo Proceed?
choice /c yn

if errorlevel 2 (
set "rcount=1"
echo list cleared
pause
goto loop
) else (
> %reqdata% echo.
> %reqquantity% echo.
for /l %%x in (1,1,!rcount!) do (
echo !rdata[%%x] >> %reqdata%
echo !rquantity[%%x] >> %reqquantity%
)

echo word file in progress...
py main.py :: put path of the main.py file 
echo DONE!
pause
)

endlocal
