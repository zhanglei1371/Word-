@echo off
copy %cd%\YYGWordAddin.dotm %USERPROFILE%\AppData\Roaming\Microsoft\Word\STARTUP >nul 2>nul
copy %cd%\YYGWordAddin.dotm %USERPROFILE%\AppData\Roaming\kingsoft\wps\startup >nul 2>nul
WordAddin.exe /regserver
echo.word插件已成功注册！
pause