@echo off
setlocal EnableDelayedExpansion

set "mes_anterior=%date:~3,2%"

if "%mes_anterior:~0,1%"=="0" (
    set "mes_anterior=%mes_anterior:~1,1%"
)

set /a "mes_anterior-=1"

if %mes_anterior%==0 (
    set "ano=!date:~-4!"
    set /a "ano-=1"
    set "mes_anterior=12"
) else if %mes_anterior%<10 (
    set "mes_anterior=0%mes_anterior%"
) 

set "pasta=C:\Users\gabriel.ribeiro\Downloads"

for %%a in ("%pasta%\*.*") do (
    set "data_arquivo=%%~ta"
    set "mes_arquivo=!data_arquivo:~3,2!"

    if "%mes_arquivo:~0,1%"=="0" (
        set "mes_arquivo=!mes_arquivo:~1,1!"
    )

    if !mes_arquivo! == %mes_anterior% (
        del "%%a"
    )
)
