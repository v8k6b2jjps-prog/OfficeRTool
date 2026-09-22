
@cls
@echo off

>nul chcp 437
setLocal EnableExtensions EnableDelayedExpansion
  
title WMI SOFT REPAIR

:: Based on Guide from
:: https://woshub.com/wmi-troubleshooting/

echo:
echo * Make sure to run as administrator
echo * Please disable any AV before check
echo:

echo Press any key to Continue
echo: & echo:

pause>nul
cls 
echo:

>nul 2>nul where powershell && (
  2>nul powershell -nop -c "(gcim Win32_Processor).AddressWidth"|>nul findstr /i /r "^32 ^64" && (
	  echo -PS- SUCCEED TO VALIDATE VALUE
  ) || (
	  set WMI_FAILUE=TRUE
	  echo -PS- FAIL TO VALIDATE VALUE
)) || (
  echo -PS- TOOL NOT EXIST
)

>nul 2>nul where WMIC && (
  2>nul WMIC path Win32_Processor get AddressWidth /format|more +1|>nul findstr /i /r "^32 ^64" && (
    echo WMIC SUCCEED TO VALIDATE VALUE
  ) || (
	set WMI_FAILUE=TRUE
	echo WMIC FAIL TO VALIDATE VALUE
)) || (
  echo WMIC TOOL NOT EXIST
)

echo:
if defined WMI_FAILUE (
  echo *** WMI STATUS = FAIL ***
) else (
  echo *** WMI STATUS = OK ***
)

echo:
echo:
 
CHOICE /C CS /M "Press C for Continue, S for Stop."
if %errorlevel% EQU 1 (
  goto :Begin
)
if %errorlevel% EQU 2 (
 timeout 4
 goto :eof
)

:Begin
echo:
winmgmt /verifyrepository
Winmgmt /salvagerepository
net stop Winmgmt /y
net start Winmgmt /y
sc config winmgmt start= disabled
net stop winmgmt /y

for %%$ in (
%windir%\system32\wbem,
%windir%\SysWOW64\wbem) do (
if exist %%$ (
  pushd %%$ && (
    for /f %%s in ('dir /b *.dll') do regsvr32 /s %%s
    2>nul wmiprvse /regserver
    sc config winmgmt start= auto
    net start winmgmt
    for /f %%s in ('dir /b *.mof') do >nul mofcomp %%s
    for /f %%s in ('dir /b *.mfl') do >nul mofcomp %%s
)))

echo:
echo * must restart the system to apply changes *
echo:
pause