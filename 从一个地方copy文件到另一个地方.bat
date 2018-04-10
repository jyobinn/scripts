@echo off&setlocal enabledelayedexpansion 
rem 这是地址
set url=C:\Users\jyo\Desktop\oldcode\
rd/s/q  !url!
mkdir  !url!
(for /f "tokens=*" %%i in (更改的文件.txt)do ( 
echo %%i
  set "dir=%%~i" 

  set "dir=!dir:~,-1!y" 

  for %%j in ("!dir!")do (
 set "h=%%~dpj"
 rem echo %%i 
rem echo !h!
set "g=!h:E:\CITS\Workspace\oraclePJ2013\=%url%!"
rem echo !g!
mkdir   !g!
copy  %%i   !g!
  )
  ))

pause