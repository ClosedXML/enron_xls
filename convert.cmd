@echo off

REM This script is used to convert batches of xlsx files, because conversion process is
REM really slow (12s per file) and basically locks the computer in the meantime.

REM Delayed expansion enables variable count to be evaluated during runtime in the loop
REM the variable is delayed expanded, when it's in exclamation marks, e.g. !COUNT!
setlocal enabledelayedexpansion
set COUNT=0
REM directory must be *absolute* path, otherwise 
set FILES_PATTERN=c:\Users\havli\source\repos\enron_xls\edrm\*.xls

REM how many files skip before taking the files
set SKIP=300
REM How many files to convert
set TAKE=100

set /a END=SKIP+TAKE

for %%F in (%FILES_PATTERN%) do (
   if !COUNT! lss %SKIP% (
       REM Ignore skipped files   
   ) else if !COUNT! lss %END% (
        echo Convert %%F		
        REM You can perform actions on the file here
		"c:\Program Files\Microsoft Office\root\Office16\excelcnv.exe" -oice "%%F" "%%F.xlsx"
    ) else (
	    echo End
        exit /b
    )
	
	set /a COUNT+=1
)