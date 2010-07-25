@echo off
cls

title PDFExec

set app_dir=%~dp0
set args=
:next_param
if [%1]==[] goto done_params
set args=%args% %1
:shift_param
shift
goto next_param
:done_params

if NOT EXIST "%app_dir%pdfexec.vbs" goto no_vbs
cscript //nologo "%app_dir%pdfexec.vbs" %args%
goto close

:no_vbs
echo The required VBScript file was not located in the application directory:
echo.
echo     %app_dir%pdfexec.vbs
goto pause

:pause
echo.
echo [press return]
pause > nul
:close
