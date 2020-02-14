@echo off

set targetPrePath=
set /p targetPrePath=[Must]Please input full path of target folder(support drag/drop):

echo %targetPrePath%

set targetCurrentPath=
set /p targetCurrentPath=[Optional]Please input output path(press Enter to use default):

echo %targetCurrentPath%

node index.js %targetPrePath% %targetCurrentPath%

pause