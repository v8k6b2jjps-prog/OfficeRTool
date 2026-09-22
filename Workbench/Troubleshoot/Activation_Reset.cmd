@cls
@echo off

echo.
echo ##### Gain Admin Privileges

rem Run as administrator, AveYo: ps\VBS version
>nul fltmc || ( set "_=call "%~dpfx0" %*"
	powershell -nop -c start cmd -args '/d/x/r',$env:_ -verb runas || (
	mshta vbscript:execute^("createobject(""shell.application"").shellexecute(""cmd"",""/d/x/r "" &createobject(""WScript.Shell"").Environment(""PROCESS"")(""_""),,""runas"",1)(window.close)"^))|| (
	cls & echo:& echo Script elavation failed& pause)
	exit )

echo.
echo ##### Running :: slmgr.vbs /rilc
echo.
1>nul 2>&1 net stop sppsvc
1>nul 2>&1 pushd "%Systemdrive%\Windows\ServiceProfiles\NetworkService\AppData\Roaming\Microsoft\SoftwareProtectionPlatform" && (
	if exist tokens.bar >nul del /q tokens.bar
	if exist tokens.dat >nul ren tokens.dat tokens.bar
)
1>nul 2>&1 pushd "%Systemdrive%\Windows\System32\spp\store\" && (
	if exist tokens.bar >nul del /q tokens.bar
	if exist tokens.dat >nul ren tokens.dat tokens.bar
	popd
)
1>nul 2>&1 pushd "%Systemdrive%\Windows\System32\spp\store\2.0\" && (
	if exist tokens.bar >nul del /q tokens.bar
	if exist tokens.dat >nul ren tokens.dat tokens.bar
	popd
)
1>nul 2>&1 net start sppsvc
cscript.exe %windir%\system32\slmgr.vbs /rilc

echo:
pause