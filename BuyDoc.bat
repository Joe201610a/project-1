@echo off
setlocal enabledelayedexpansion

set "names=Item-names.txt"
set "reqdata=requested-data.txt"
set "reqprices=requested-prices.txt"
set "reqquantity=requested-quantity.txt"

set "count=0"

for /f "delims=" %%a in (!names!) do (
    set /a count+=1
    set "name[!count!]=%%a"
)

set /p customername="Customer name: "

set "rcount=1"
set "totalprice=0"

:loop
    cls
    echo - choose from below - (press 0 to finish)
    echo.

    for /l %%x in (1,1,!count!) do (
        echo %%x - !name[%%x]!
    )

    set /p r=""

    if !r! == 0 goto next
    if !r! gtr !count! (
        echo This option isnt available...
        pause
        goto loop
    )
    if !r! lss 0 (
        echo This option isnt available...
        pause
        goto loop
    )
    echo Unit price:
    set /p p=

    echo Quantity:
    set /p q=""

    set "rdata[!rcount!]=!name[%r%]!"
    set "rprice[!rcount!]=!p!"
    set "rquantity[!rcount!]=!q!"
    set /a totalprice += !p! * !q!
    set /a rcount+=1
    goto loop

:next

set /a rcount-=1
cls
echo Requested data - name : unit price : quantity
echo.
for /l %%x in (1,1,!rcount!) do (
    set "temp=!rdata[%%x]!"
    echo %%x - !rdata[%%x]! : !rprice[%%x]!$ : !rquantity[%%x]!
)
echo.
echo Total price = !totalprice!$
echo.
echo Proceed?
choice /c yn

if errorlevel 2 (
    set "rcount=1"
    set "totalprice=0"
    echo list cleared
    pause
    goto loop
) else (
    > %reqdata% echo.
    > %reqquantity% echo.
    > %reqprices% echo.
    for /l %%x in (1,1,!rcount!) do (
        echo !rdata[%%x]! >> %reqdata%
        echo !rprice[%%x]! >> %reqprices%
        echo !rquantity[%%x]! >> %reqquantity%
    )
    echo !customername! >> %reqdata%

    echo word file in progress...
    py main.py
    echo DONE!
    pause
)

endlocal
