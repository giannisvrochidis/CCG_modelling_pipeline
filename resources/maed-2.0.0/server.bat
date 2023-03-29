echo off

set PORT=%1
if "" == "%1" set /p PORT="Please specify a port (ex. 8765) to start the MAED server: "

echo Stopping any process running on the specified port.
FOR /f "tokens=5" %%T IN ('netstat -ano ^| findstr %PORT% ') DO ( set /a PID=%%T )
taskkill /f /pid %PID% > nul 2> nul

echo Starting MAED server.
%~dp0\php\win32\php-7.4.7\php -S localhost:%PORT% -t %~dp0\project
