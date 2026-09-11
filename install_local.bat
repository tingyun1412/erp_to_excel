@echo off
setlocal
set "SRC=%~dp0"
set "DEST=%LOCALAPPDATA%\ShippingAutomationTool"

echo Installing to local machine, please wait...
echo Source: %SRC%
echo Destination: %DEST%
echo.

if not exist "%DEST%" mkdir "%DEST%"

robocopy "%SRC%app" "%DEST%\app" /E /XO /R:2 /W:2 /NFL /NDL /NJH /NJS
robocopy "%SRC%python" "%DEST%\python" /E /XO /R:2 /W:2 /NFL /NDL /NJH /NJS
robocopy "%SRC%ms-playwright" "%DEST%\ms-playwright" /E /XO /R:2 /W:2 /NFL /NDL /NJH /NJS

set "VBS=%TEMP%\make_shortcut_tmp.vbs"
> "%VBS%" echo Set oWS = WScript.CreateObject("WScript.Shell")
>> "%VBS%" echo sLinkFile = oWS.SpecialFolders("Desktop") ^& "\ShippingTool.lnk"
>> "%VBS%" echo Set oLink = oWS.CreateShortcut(sLinkFile)
>> "%VBS%" echo oLink.TargetPath = "%DEST%\python\pythonw.exe"
>> "%VBS%" echo oLink.Arguments = """%DEST%\app\webview_launcher.py"""
>> "%VBS%" echo oLink.WorkingDirectory = "%DEST%\app"
>> "%VBS%" echo oLink.IconLocation = "shell32.dll,166"
>> "%VBS%" echo oLink.WindowStyle = 1
>> "%VBS%" echo oLink.Save
cscript //nologo "%VBS%"
del "%VBS%"

attrib +h "%DEST%\app" 2>nul
attrib +h "%DEST%\python" 2>nul
attrib +h "%DEST%\ms-playwright" 2>nul

echo.
echo Done. A shortcut named "ShippingTool" has been created on your Desktop.
echo Run this installer again anytime to sync the latest updates.
echo.
pause
