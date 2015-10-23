@echo off

for /F "tokens=1* delims= " %%A in ('date /T') do set CDATE=%%B
for /F "tokens=1,2 eol=/ delims=/ " %%A in ('date /T') do set mm=%%B
for /F "tokens=1,2 delims=/ eol=/" %%A in ('echo %CDATE%') do set dd=%%B
for /F "tokens=2,3 delims=/ " %%A in ('echo %CDATE%') do set yyyy=%%B
set dateNew=%mm%-%dd%-%yyyy%

set month-num=%date:~3,2%
if %mm%==01 set mo-name=January
if %mm%==02 set mo-name=February
if %mm%==03 set mo-name=March
if %mm%==04 set mo-name=April
if %mm%==05 set mo-name=May
if %mm%==06 set mo-name=June
if %mm%==07 set mo-name=July
if %mm%==08 set mo-name=August
if %mm%==09 set mo-name=September
if %mm%==10 set mo-name=October
if %mm%==11 set mo-name=November
if %mm%==12 set mo-name=December

set path=D:\Personal\Journal\%yyyy%\%mo-name%\%dateNew%.docx
set monthPath=D:\Personal\Journal\%yyyy%\%mo-name%
set yearPath=D\Personal\Journal\%yyyy%

if not exist "%yearPath%" mkdir %yearPath%
if not exist "%monthPath%" mkdir %monthPath%
if not exist "%path%" copy /y nul "%path%"
start %path%