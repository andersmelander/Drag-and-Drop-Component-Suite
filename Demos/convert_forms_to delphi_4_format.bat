@echo off
cls
echo Drag and Drop Component Suite v4
echo +-----------------------------------------------------------
echo :
echo : This will convert the forms of all the demo projects from
echo : text format (compatible with Delphi 5 and 6) to binary
echo : format (compatible with Delphi 4 and 5).
echo :
echo : If you haven't already done so, please make sure you read 
echo : the note about "Property does not exist" errors in the
echo : installation guide in the readme.txt file.
echo :
echo +-----------------------------------------------------------
pause

echo Searching for CONVERT.EXE...
for %%v in (%path%) do if exist %%v\convert.exe goto use_path
if exist %1\CONVERT.EXE goto use_param
goto no_convert

:use_path
set cp=
echo CONVERT.EXE found in path:
for %%v in (%path%) do if exist %%v\convert.exe echo %%v
goto do_it

:use_param
set cp="%1\"
echo CONVERT.EXE found at specified location:
echo %1
goto do_it

:no_convert
echo +-----------------------------------------------------------
echo :
echo : Error:
echo : The Delphi Form Conversion Utility CONVERT.EXE was
echo : not found in your path.
echo :
echo : You can either modify your path to include the
echo : directory where CONVERT.EXE is located or specify
echo : the path to CONVERT.EXE as a parameter to this
echo : batch file like this:
echo :
echo :  convert_forms_blah_blah "c:\program files\borland\delphi 4\bin"
echo :
echo +-----------------------------------------------------------
pause
goto exit

:do_it
set _System=
for %%v in (%path%) do if exist %%v\xcopy.* goto xcopy_found
if "%SystemRoot%"=="" set _System=%WinDir%\System\
if not "%SystemRoot%"=="" set _System=%SystemRoot%\System32\

:xcopy_found
echo Copying forms...
%_System%xcopy /s *.dfm ..\D4Forms\*.txt>nul
if exist ..\D6Forms\*.* goto dont_backup
echo Backing up forms...
%_System%xcopy /s *.dfm ..\D6Forms\*.*>nul
echo This directory contains a backup copy of the original form files in Delphi 5/6 format.>..\D6Forms\readme.txt
echo ->>..\D6Forms\readme.txt
echo Run the batch file restore_forms_to_delphi_6_format.bat located in this directory>>..\D6Forms\readme.txt
echo to restore the demo form files to Delphi 5/6 format.>>..\D6Forms\readme.txt

echo @echo This batch file will restore all the demo form files to Delphi 5/6 format.>..\D6Forms\restore_forms_to_delphi_6_format.bat
echo @%_System%xcopy /s *.dfm ..\demo\*.*>>..\D6Forms\restore_forms_to_delphi_6_format.bat
echo @pause>>..\D6Forms\restore_forms_to_delphi_6_format.bat
:dont_backup

echo Creating list of forms...
dir /s /b ..\D4Forms\*.txt>.\demo_forms.txt
echo Converting forms...
%cp%convert @demo_forms.txt>nul
echo Copying converted forms...
%_System%xcopy /s ..\D4Forms\*.dfm .\*.*>nul
echo Cleaning up...
del /s /q ..\D4Forms\*.*>nul
rd /s /q ..\D4Forms>nul
del .\demo_forms.txt>nul
echo Done!
echo ---------------------------------------------------------
pause

:exit
set _System=
set cp=
