@echo off

for /F "tokens=1* delims= " %%A in ('date /T') do set CDATE=%%B
for /F "tokens=1,2 eol=/ delims=/ " %%A in ('date /T') do set mm=%%B
for /F "tokens=1,2 delims=/ eol=/" %%A in ('echo %CDATE%') do set dd=%%B
for /F "tokens=2,3 delims=/ " %%A in ('echo %CDATE%') do set yyyy=%%B
set dateNew=%mm%-%dd%-%yyyy%

set month-num=%date:~3,2%
if %mm%==01 set mo-name=01-January
if %mm%==02 set mo-name=02-February
if %mm%==03 set mo-name=03-March
if %mm%==04 set mo-name=04-April
if %mm%==05 set mo-name=05-May
if %mm%==06 set mo-name=06-June
if %mm%==07 set mo-name=07-July
if %mm%==08 set mo-name=08-August
if %mm%==09 set mo-name=09-September
if %mm%==10 set mo-name=10-October
if %mm%==11 set mo-name=11-November
if %mm%==12 set mo-name=12-December

set path=C:\Users\nickg\Documents\Personal\Journal\JournalFiles\%yyyy%\%mo-name%\%dateNew%.docx
set monthPath=C:\Users\nickg\Documents\Personal\Journal\JournalFiles\%yyyy%\%mo-name%
set yearPath=C:\Users\nickg\Documents\Personal\Journal\JournalFiles\%yyyy%

if not exist "%yearPath%" mkdir %yearPath%
if not exist "%monthPath%" mkdir %monthPath%
if not exist "%path%" copy /y nul "%path%"
start %path%