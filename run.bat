@echo off

set targetPrePath=
set /p targetPrePath=Please input full-path(parent of the target folder):

echo %targetPrePath%

set targetCurrentPath=
set /p targetCurrentPath=Please input target folder name(name of target folder, not path):

echo %targetCurrentPath%

node index.js %targetPrePath% %targetCurrentPath%

pause