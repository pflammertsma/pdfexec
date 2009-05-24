@echo off

:: Set up some default variables
set version=1.3
set app_dir=%~dp0
set ini_file=%app_dir%lhsexec.ini
set tmp_file=%app_dir%~tmp_lhs.bat
set param_1=
set param_2=
set no_close=0
set ignore_errors=0
set no_open=0
set pause=0
set pause_errors=0
set verbose=0
set style=
set no_latex=0

echo LHSEXEC v%version%
echo Wrapper for lhs2TeX and LaTeX, copyright (c) 2008, Paul Lammertsma
echo http://paul.luminos.nl
echo For usage and command line switches, start using LHSEXEC /?
echo.

:: Check an individual parameter and return
:next_param
if [%1]==[] goto done_params
if [%1]==[/?] goto help
if [%1]==[-?] goto help
if [%1]==[/h] goto help
if [%1]==[-h] goto help
if [%1]==[/H] goto help
if [%1]==[-H] goto help
if [%1]==[/c] goto no_close
if [%1]==[-c] goto no_close
if [%1]==[/C] goto no_close
if [%1]==[-C] goto no_close
if [%1]==[/i] goto ignore_errors
if [%1]==[-i] goto ignore_errors
if [%1]==[/I] goto ignore_errors
if [%1]==[-I] goto ignore_errors
if [%1]==[/o] goto no_open
if [%1]==[-o] goto no_open
if [%1]==[/O] goto no_open
if [%1]==[-O] goto no_open
if [%1]==[/pe] goto pause_errors
if [%1]==[-pe] goto pause_errors
if [%1]==[/PE] goto pause_errors
if [%1]==[-PE] goto pause_errors
if [%1]==[/p] goto pause
if [%1]==[-p] goto pause
if [%1]==[/P] goto pause
if [%1]==[-P] goto pause
if [%1]==[/s] goto style
if [%1]==[-s] goto style
if [%1]==[/S] goto style
if [%1]==[-S] goto style
if [%1]==[/v] goto verbose
if [%1]==[-v] goto verbose
if [%1]==[/V] goto verbose
if [%1]==[-V] goto verbose
if [%param_1%]==[] goto param_1
if [%param_2%]==[] goto param_2
goto shift_param

:: Set the various parameters
:no_close
set no_close=1
goto :shift_param
:ignore_errors
set ignore_errors=1
goto :shift_param
:no_open
set no_open=1
goto :shift_param
:pause
set pause=1
goto :shift_param
:pause_errors
set pause_errors=1
goto :shift_param
:verbose
set verbose=1
goto :shift_param
:style
if [%2]==[verb] goto style_ok
if [%2]==[tt] goto style_ok
if [%2]==[math] goto style_ok
if [%2]==[poly] goto style_ok
if [%2]==[code] goto style_no_latex
if [%2]==[newcode] goto style_no_latex
echo Unsupported style "%2" requested; parameter dropped.
goto shift_param
:style_no_latex
set no_latex=1
:style_ok
set style=%2
goto :shift_param
:param_1
set param_1=%1
goto :shift_param
:param_2
set param_2=%1
goto :shift_param

:: Shift the next parameter up one position
:shift_param
shift
goto next_param

:: Check if parameters were sent
:done_params
if [%param_1%]==[] goto ask_file1
goto check_file1

:: If not, ask
:fail_file1
if NOT EXIST "%param_1%.lhs" goto fail_file1b
set param_1=%param_1%.lhs
goto check_file1
:fail_file1b
echo The specified input file could not be found:
echo   "%param_1%"
:ask_file1
echo Please specify the input filename:
set /p param_1=">"
if [%param_1%]==[] goto abort
echo.
:check_file1
if NOT EXIST "%param_1%" goto fail_file1
if [%param_2%]==[] goto ask_file2
goto start
:ask_file2
echo Please specify the output filename, excluding extension:
set /p param_2=">"
if [%param_2%]==[] goto abort
echo.

:: All is okay; we can continue
:start

:: Grab configuration information
if NOT EXIST "%ini_file%" goto config_missing
copy "%ini_file%" "%tmp_file%" > nul
if NOT EXIST "%tmp_file%" goto config_error
call "%tmp_file%" > nul
del "%tmp_file%" > nul

:: Check if the configuration is OK

if ["%lhs2tex_directory%"]==[""] goto config_missing
if ["%latex_directory%"]==[""] goto config_missing
if NOT EXIST "%lhs2tex_directory%\lhs2TeX.exe" goto no_lhs2tex
if NOT EXIST "%latex_directory%\pdflatex.exe" goto no_latex

echo Input:           %param_1%
if [%style%]==[] goto info_no_style
echo Style:           %style%
goto info_done
:info_no_style
echo Style:           (default)
:info_done
echo Output TeX file: %param_2%.tex
echo Output PDF file: %param_2%.pdf
echo.
echo Generating TeX files...
if [%style%]==[] goto no_style
"%lhs2tex_directory%\lhs2TeX.exe" --%style% %param_1% > %param_2%.tex
goto lhs2tex_done
:no_style
"%lhs2tex_directory%\lhs2TeX.exe" %param_1% > %param_2%.tex
:lhs2tex_done
if NOT [%ignore_errors%]==[1] if [%ERRORLEVEL%]==[1] goto error

if [%no_latex%]==[1] goto latex_done
if [%no_close%]==[1] goto create
if NOT EXIST "%param_2%.pdf" goto create
echo.
echo Closing PDF...
"%latex_directory%\pdfclose.exe" --file=%param_2%.pdf 2>nul
taskkill /FI "WINDOWTITLE eq %param_2%.pdf - Adobe Acrobat Professional" > nul
taskkill /FI "WINDOWTITLE eq %param_2%.pdf - Adobe Acrobat" > nul
taskkill /FI "WINDOWTITLE eq %param_2%.pdf - Adobe Reader" > nul

:create
echo.
echo Generating PDF...
if NOT [%verbose%]==[1] goto latex_quiet
"%latex_directory%\pdflatex.exe" %param_2%.tex
goto latex_done
:latex_quiet
"%latex_directory%\pdflatex.exe" -quiet %param_2%.tex
:latex_done
echo.
if NOT [%ignore_errors%]==[1] if [%ERRORLEVEL%]==[1] goto error
if NOT [%no_open%]==[1] "%latex_directory%\pdfopen.exe" --file=%param_2%.pdf
if NOT [%ignore_errors%]==[1] if [%ERRORLEVEL%]==[1] goto error
echo Done.
goto end

:help
echo Compiles TeX files from .lhs Haskell files using lhs2TeX and outputs a PDF.
echo.
echo LHSEXEC input_file output_name [switches]
echo.
echo   input_file     Literal Haskell file to input into lhs2TeX
echo   output_name    Filename, excluding file extension, to use as filename for
echo                  .tex file and .pdf file
echo.
echo The list of switches is as follows:
echo.
echo.  /?             This help screen
echo   /s style       Set the lhs2TeX style parameter to one of the following:
echo                    verb      verbatim text
echo                    tt        formatted verbatim text
echo                    math      mathematical formatting
echo                    poly      aligned mathematical formatting
echo                    code      remove comments
echo                    newcode   remove comments, formatted
echo   /i             Ignore any errors and continue (i.e. resume batch)
echo   /o             Do not automatically open Acrobat
echo   /c             Do not automatically close Acrobat
echo   /p             Require keystroke to terminate
echo   /pe            Require keystroke to terminate, only on errors
echo   /v             Verbose mode for LaTeX
goto end

:: Various error messages

:no_lhs2tex
echo lhs2TeX could not be found in the following directory:
echo   "%lhs2tex_directory%"
echo.
echo LHSEXEC expected to find the following mandatory lhs2TeX executable:
echo   "%lhs2tex_directory%\lhs2TeX.exe"
goto review_config
:no_latex
echo LaTeX could not be found in the following directory:
echo   "%latex_directory%"
echo.
echo LHSEXEC expected to find the following mandatory LaTeX executable:
echo   "%latex_directory%\pdflatex.exe"
goto review_config
:config_error
echo It appears that the configuration file cannot be accessed.
echo Please ensure that you have write access to the LHSEXEC directory. This file
echo can be found here:
echo.
echo   %ini_file%
goto end
:config_missing
echo It appears the configuration file is missing or invalid. This file can be found
echo here:
echo.
echo   %ini_file%
echo.
echo LHSEXEC can restore the default values.
echo.
set /p create=Restore default values (Y/N)?
if [%create%]==[Y] goto create_config
if [%create%]==[y] goto create_config
goto end
:create_config
echo set latex_directory=C:\Program Files\lhs2tex > "%ini_file%"
echo set lhs2tex_directory=C:\Program Files\MikTeX\miktex\bin >> "%ini_file%"
echo.
echo The configuration file has been created containing default values. Please
echo review it to specify the correct directories for lhs2TeX and LaTeX.
echo.
echo This file can be found here:
echo.
echo   %ini_file%
goto end
:review_config
echo.
echo Please review the configuration file to specify these directories.
echo.
echo This file can be found here:
echo.
echo   %ini_file%
goto end
:error
echo.
echo One or more errors occurred!
echo Please review the messages above.
goto end

:: Finish

:abort
echo.
echo Empty query returned; aborting.
:end
if [%pause%]==[1] goto pause
if [%ERRORLEVEL%]==[0] goto close
if [%ignore_errors%]==[0] goto pause
if [%pause_errors%]==[1] goto pause_error
goto close
:pause_error
echo.
echo One or more errors occurred!
echo Please review the messages above.
:pause
echo.
echo Press any key to continue.
pause > nul
:close
