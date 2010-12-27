:: Running demosync from command line (cmd)
:: Testet on Windows XP with WAMP installed
:: Output is put in the logs folders

:: Add the following as a scheduled task:
:: C:\wamp\www\jm-bookingtest\libs\ExchangePHP\demo\demosync.bat

@echo off
For /f "tokens=1-3 delims=/. " %%a in ('date /t') do (set mydate=%%c-%%b-%%a)
For /f "tokens=1-2 delims=/:" %%a in ("%TIME%") do (set mytime=%%a%%b)

C:\wamp\bin\php\php5.2.6\php.exe c:\wamp\www\jm-bookingtest\libs\ExchangePHP\demo\index.php > "c:\wamp\www\jm-bookingtest\libs\ExchangePHP\demo\logs\log-%mydate%_%mytime%.txt"
