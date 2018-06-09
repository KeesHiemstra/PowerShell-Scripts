@echo off
::
:::::::::::::::::::::::::::::::::::::::
:: by idris.hassini@hp.com           ::
:: SFS & QM @ Sara Lee International ::
:: 2009/07/16                        :: 
:::::::::::::::::::::::::::::::::::::::
::
set _wrkdir=C:\rpmtools\ownrPage\
set _input=%_wrkdir%input.txt
set _QM_RPRT_Name=foldersize.txt
::
for /F "eol=# tokens=2 delims=," %%i in (%_input%) do (
 %_wrkdir%echo64 -n .
 for /F "tokens=1 delims=:" %%j in ("%%i") do (
  %_wrkdir%echo64 -n .
  for /F %%k in ('dir /b %%i') do (
   %_wrkdir%echo64 -n .
   echo Size report for %%k > %%i\%%k\%_QM_RPRT_Name%
   echo. >> %%i\%%k\%_QM_RPRT_Name%
   %%j:
   cd\
   cd %%i\%%k
   %_wrkdir%diruse /m /* . >> %%i\%%k\%_QM_RPRT_Name%
   echo ===================== >> %%i\%%k\%_QM_RPRT_Name%
  )
 )
)