@echo off
WordAddin.exe /Unregserver
echo.����ǿ�ƹر�����word�ĵ������ȱ��棬���ֶ��رգ� & pause
taskkill /f /im winword.exe >nul 2>nul
del %USERPROFILE%\AppData\Roaming\Microsoft\Word\STARTUP\YYGWordAddin.dotm >nul 2>nul
del %USERPROFILE%\AppData\Roaming\kingsoft\wps\startup\YYGWordAddin.dotm >nul 2>nul
echo.����ѳɹ�ж�أ�
pause