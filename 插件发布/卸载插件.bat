@echo off
WordAddin.exe /Unregserver
echo.即将强制关闭所有word文档，请先保存，并手动关闭！ & pause
taskkill /f /im winword.exe >nul 2>nul
del %USERPROFILE%\AppData\Roaming\Microsoft\Word\STARTUP\YYGWordAddin.dotm >nul 2>nul
del %USERPROFILE%\AppData\Roaming\kingsoft\wps\startup\YYGWordAddin.dotm >nul 2>nul
echo.插件已成功卸载！
pause