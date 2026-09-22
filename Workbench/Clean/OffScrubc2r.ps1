#requires -version 5.0
#requires -runasadministrator

<#
version      : 5.0
Release Date : 08.02.2024

- based on OffScrubc2r.vbs
- Re-written by Dark Dinosaur [MDL]
- Some code implements different,
  than the original VBS file

Set-Location "HKCU:\"
RI           "HKCU:\XXX\XXX\XXX"

Set-Location "HKLM:\"
GPV          "HKLM:\XXX\XXX\XXX"

Set-Location "$($env:SystemDrive)\"
GI           "$($env:WINDIR)\explorer.exe"

GI,GP,GPV -- RI,RP -- SI, SP
[ Push-Location ] -> [ `gi .` / `GCI .` ] -> [ XXX.Property | [ `gci $_` / `GPV $_` ] ]

$hDefKey = "HKCU"
$sSubKeyName = "SOFTWARE\Microsoft\AppV\ISV"
Set-Location "HKCU:\"
Push-Location "$($hDefKey):$($sSubKeyName)" -ea 0

if (@(Get-Location).Path -ne 'HKCU:\') {
  DO }
#>

#cls
Write-Host

$dicKeepSku = @{}
$Start_Time = $(Get-Date -Format hh:mm:ss)
$IgnoreCase = [Text.RegularExpressions.RegexOptions]::IgnoreCase

Set-Location "HKLM:\"
$sPackageGuid = $null

@("SOFTWARE\Microsoft\Office\15.0\ClickToRun",
  "SOFTWARE\Microsoft\Office\16.0\ClickToRun",
  "SOFTWARE\Microsoft\Office\ClickToRun" ) | % {
    try {
      $sPackageGuid = gpv $_ PackageGUID -ea 0
    } catch{}}

Function IsC2R {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $Value,
	 
	 [parameter(Mandatory=$false)]
     [bool] $FastSearch
   )

  $OREF          = "^(.*)(\\ROOT\\OFFICE1)(.*)$"
  $MSOFFICE      = "^(.*)(\\Microsoft Office)(.*)$"
  $OREFROOT      = "^(.*)(Microsoft Office\\root\\)(.*)$"
  $OCOMMON	     = "^(.*)(\\microsoft shared\\ClickToRun)(.*)$"

  
  if (($FastSearch -ne $null) -and ($FastSearch -eq $true)) {
	if ([REGEX]::IsMatch(
      $Value,$MSOFFICE,$IgnoreCase)) {
        return $true }
	return $false
  }
  
  if ([REGEX]::IsMatch(
    $Value,$OREF,$IgnoreCase)) {
      return $true }
  if ([REGEX]::IsMatch(
    $Value,$MSOFFICE,$IgnoreCase)) {
      return $true }
  if ([REGEX]::IsMatch(
    $Value,$OREFROOT,$IgnoreCase)) {
      return $true }
  if ([REGEX]::IsMatch(
    $Value,$OCOMMON,$IgnoreCase)) {
      return $true }
             
  return $false
}
Function GetExpandedGuid {
    param (
        [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
        [ValidatePattern('^[0-9a-fA-F]{32}$')]
        [ValidateScript( { [System.Guid]::Parse($_) -is [System.Guid] })]
        [string] $sGuid
    )

    if (($sGuid.Length -ne 32) -or (
        $sGuid -notmatch '00F01FEC')) {
        return $null }

    $output = ""
    ([ordered]@{
    1=$sGuid.ToCharArray(0,8)
    2=$sGuid.ToCharArray(8,4)
    3=$sGuid.ToCharArray(12,4)}).GetEnumerator() | % {
        [ARRAY]::Reverse($_.Value)
        $output += (-join $_.Value) }
    $sArr = $sGuid.ToCharArray()
    ([ordered]@{
    17=20
    21=32 }).GetEnumerator() | % {
    $_.Key..$_.Value | % {
        if ($_ % 2) {
        $output += $sArr[$_]
    } else {
        $output += $sArr[$_-2] }} }
    return [System.Guid]::Parse(
        -join $output).ToString().ToUpper()
}

Function CheckDelete {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $sProductCode
   )

   # FOR GUID FORMAT
   # {90160000-008C-0000-1000-0000000FF1CE}

   # ensure valid GUID length
   if ($sProductCode.Length -ne 38) {
     return $false }	

    # only care if it's in the expected ProductCode pattern
	if (-not(
	  InScope $sProductCode)) {
        return $false }
	
    # check if it's a known product that should be kept
    if ($dicKeepSku.ContainsKey($sProductCode)) {
      return $false }
	
  return $True
}
Function InScope {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $sProductCode
   )
   
   $PRODLEN = 13
   $OFFICEID = "0000000FF1CE}"
   if ($sProductCode.Length -ne 38) {
     return $false }

   $sProd = $sProductCode.ToUpper()
   if ($sProd.Substring($sProd.Length-13,$PRODLEN) -ne $OFFICEID ) {
     if ($sPackageGuid -and ($sProd -eq $sPackageGuid.ToUpper())) {
       return $True }
     switch ($sProductCode)
     {
       "{6C1ADE97-24E1-4AE4-AEDD-86D3A209CE60}" {return $True}
       "{9520DDEB-237A-41DB-AA20-F2EF2360DCEB}" {return $True}
       "{9AC08E99-230B-47e8-9721-4577B7F124EA}" {return $True}
     }
     return $false }
   
   if ([INT]$sProd.Substring(3,2) -gt 14) {
     switch ($sProd.Substring(10,4))
     {
       "007E" {return $True}
       "008F" {return $True}
       "008C" {return $True}
       "24E1" {return $True}
       "237A" {return $True}
       "00DD" {return $True}
       Default {return $false}
     }
   }

    return $false
}
function Get-Shortcut {
<#
.SYNOPSIS
    Get information about a Shortcut (.lnk file)
.DESCRIPTION
    Get information about a Shortcut (.lnk file)
.PARAMETER Path
    File
.EXAMPLE
    Get-Shortcut -Path 'C:\Portable\Test.lnk'
 
    Link : Test.lnk
    TargetPath : C:\Portable\PortableApps\Notepad++Portable\Notepad++Portable.exe
    WindowStyle : 1
    IconLocation : ,0
    Hotkey :
    Target : Notepad++Portable.exe
    Arguments :
    LinkPath : C:\Portable\Test.lnk
#>

    [CmdletBinding(ConfirmImpact='None')]
    param(
        [string] $path
    )

    begin {
        Write-Verbose -Message "Starting [$($MyInvocation.Mycommand)]"
        $obj = New-Object -ComObject WScript.Shell
    }

    process {
        if (Test-Path -Path $Path) {
            $ResolveFile = Resolve-Path -Path $Path
            if ($ResolveFile.count -gt 1) {
                Write-Error -Message "ERROR: File specification [$File] resolves to more than 1 file."
            } else {
                Write-Verbose -Message "Using file [$ResolveFile] in section [$Section], getting comments"
                $ResolveFile = Get-Item -Path $ResolveFile
                if ($ResolveFile.Extension -eq '.lnk') {
                    $link = $obj.CreateShortcut($ResolveFile.FullName)

                    $info = @{}
                    $info.Hotkey = $link.Hotkey
                    $info.TargetPath = $link.TargetPath
                    $info.LinkPath = $link.FullName
                    $info.Arguments = $link.Arguments
                    $info.Target = try {Split-Path -Path $info.TargetPath -Leaf } catch { 'n/a'}
                    $info.Link = try { Split-Path -Path $info.LinkPath -Leaf } catch { 'n/a'}
                    $info.WindowStyle = $link.WindowStyle
                    $info.IconLocation = $link.IconLocation

                    New-Object -TypeName PSObject -Property $info
                } else {
                    Write-Error -Message 'Extension is not .lnk'
                }
            }
        } else {
            Write-Error -Message "ERROR: File [$Path] does not exist"
        }
    }

    end {
        Write-Verbose -Message "Ending [$($MyInvocation.Mycommand)]"
    }
}
Function CleanShortcuts {
   param (
     [parameter(Mandatory=$True)]
     [string] $sFolder
   )

 Set-Location "c:\"

 if (-not (
   Test-Path $sFolder )) {
     return; }

 dir $sFolder -Filter *.lnk -Recurse -ea 0 | % {
    $Shortcut = Get-Shortcut(
      $_.FullName) -ea 0
    if ($Shortcut -and $Shortcut.TargetPath -and (
      $Shortcut.TargetPath|IsC2R)) {
          RI $_.FullName -Force -ea 0  }}
}
function UninstallOfficeC2R {
$URL = 
  "http://officecdn.microsoft.com/pr/wsus/setup.exe"

$Path = 
  "$env:WINDIR\temp\setup.exe"

$XML = 
  "$env:WINDIR\temp\RemoveAll.xml"


$CODE = @"
<Configuration> 
  <Remove All="TRUE"> 
</Remove> 
  <Display Level="None" AcceptEULA="TRUE" />   
</Configuration>
"@

try {
  "*** -- build the remove.xml"
  $CODE | Out-File $XML
  "*** -- ODT not available. Try to download"
  (New-Object System.Net.WebClient).DownloadFile($URL, $Path)
}
catch { }

Set-Location "$env:SystemDrive\"
Push-Location "$env:WINDIR\temp\"
if ([IO.FILE]::Exists(
  $Path)) {
    $Proc = start $Path -arg "/configure RemoveAll.xml" -Wait -WindowStyle Hidden -PassThru -ea 0
    "*** -- ODT uninstall for OfficeC2R returned with value:$($Proc.ExitCode)" }

if ($Proc -and $Proc.ExitCode -eq 0) {
  "*** -- Use unified ARP uninstall command [No-Need]"
  return }

"*** -- Use unified ARP uninstall command"

try {
  $HashList = GetUninstall }
catch {
  $HashList = $null }

$arrayList = @{}
$OfficeClickToRun = $null

if ($HashList) {
  foreach ($key in $HashList.keys) {
    $value = $HashList[$key]
    if (($value -notlike "*OfficeClickToRun.exe*") -and (
      $false -eq ($value|CheckDelete) )) {
      continue }
    $data  = $value.Split( )
    if ($data) {
      0..$data.Count | % {
        if ($data[$_] -match 'productstoremove=') {
          $data[$_] = "productstoremove=AllProducts" }}
    
    $value   = $data -join (' ')
    $value  += ' displaylevel=false'
    $prefix  = $value.Split('"')
    try {
      $OfficeClickToRun = $prefix[1]
      $arrayList.Add($key,$prefix[2]) }
    catch {}
}}}

foreach ($key in $arrayList.Keys) {
  if ([IO.FILE]::Exists($OfficeClickToRun)) {
    $value = $arrayList[$key]
    $Proc = start $OfficeClickToRun -Arg $value -Wait -WindowStyle Hidden -PassThru -ea 0
    "*** -- uninstall command: $arg, exit code value: $($Proc.ExitCode)"
}}

return
}
Function CloseOfficeApps {
$dicApps = @{}
$dicApps.Add("appvshnotify.exe","appvshnotify.exe")
$dicApps.Add("integratedoffice.exe","integratedoffice.exe")
$dicApps.Add("integrator.exe","integrator.exe")
$dicApps.Add("firstrun.exe","firstrun.exe")
$dicApps.Add("communicator.exe","communicator.exe")
$dicApps.Add("msosync.exe","msosync.exe")
$dicApps.Add("OneNoteM.exe","OneNoteM.exe")
$dicApps.Add("iexplore.exe","iexplore.exe")
$dicApps.Add("mavinject32.exe","mavinject32.exe")
$dicApps.Add("werfault.exe","werfault.exe")
$dicApps.Add("perfboost.exe","perfboost.exe")
$dicApps.Add("roamingoffice.exe","roamingoffice.exe")
$dicApps.Add("officeclicktorun.exe","officeclicktorun.exe")
$dicApps.Add("officeondemand.exe","officeondemand.exe")
$dicApps.Add("OfficeC2RClient.exe","OfficeC2RClient.exe")
$dicApps.Add("explorer.exe","explorer.exe")
$dicApps.Add("msiexec.exe","msiexec.exe")
$dicApps.Add("ose.exe","ose.exe")
$dicList = $dicApps.Values -join "|"

$Process = gwmi -Query "Select * From Win32_Process"
$Process | ? {
  [REGEX]::IsMatch($_.Name,$dicList, $IgnoreCase)} | % {
    try {($_).Terminate()|Out-Null} catch {} }

$Process = gwmi -Query "Select * From Win32_Process"
$Process | % {
  $ExecuePath = ($_).Properties | ? Name -EQ ExecutablePath | select Value
  if ($ExecuePath -and $ExecuePath.Value) {
    if ($ExecuePath.Value|IsC2R) {
        try {
          ($_).Terminate()|Out-Null}
        catch {} }}}
}
Function DelSchtasks {
SCHTASKS /Delete /F /TN "C2RAppVLoggingStart" *>$null
SCHTASKS /Delete /F /TN "FF_INTEGRATEDstreamSchedule" *>$null
SCHTASKS /Delete /F /TN "Microsoft Office 15 Sync Maintenance for {d068b555-9700-40b8-992c-f866287b06c1}" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Automatic Updates 2.0" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Automatic Updates" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office ClickToRun Service Monitor" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Feature Updates Logon" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Feature Updates" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\Office Performance Monitor" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\OfficeInventoryAgentFallBack" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\OfficeInventoryAgentLogOn" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\OfficeTelemetryAgentFallBack" *>$null
SCHTASKS /Delete /F /TN "Microsoft\Office\OfficeTelemetryAgentLogOn" *>$null
SCHTASKS /Delete /F /TN "Office 15 Subscription Heartbeat" *>$null
SCHTASKS /Delete /F /TN "Office Background Streaming" *>$null
SCHTASKS /Delete /F /TN "Office Subscription Maintenance" *>$null
}
Function ClearShellIntegrationReg {
Set-Location "HKLM:\"
RI "HKLM:SOFTWARE\Classes\Protocols\Handler\osf" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{573FFD05-2805-47C2-BCE0-5F19512BEB8D}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{8BA85C75-763B-4103-94EB-9470F12FE0F7}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{CD55129A-B1A1-438E-A425-CEBC7DC684EE}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}" -Force -ea 0 -Recurse

RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 1 (ErrorConflict)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 2 (SyncInProgress)" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\ShellIconOverlayIdentifiers\Microsoft SPFS Icon Overlay 3 (InSync)" -Force -ea 0 -Recurse
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{B28AA736-876B-46DA-B3A8-84C5E30BA492}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{8B02D659-EBBB-43D7-9BBA-52CF22C5B025}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{42042206-2D85-11D3-8CFF-005004838597}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{993BE281-6695-4BA5-8A2A-7AACBFAAB69E}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{C41662BB-1FA0-4CE0-8DC5-9B7F8279FF97}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{506F4668-F13E-4AA1-BB04-B43203AB3CC0}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{D66DC78C-4F61-447F-942B-3FB6980118CF}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{46137B78-0EC3-426D-8B89-FF7C3A458B5E}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{8BA85C75-763B-4103-94EB-9470F12FE0F7}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{CD55129A-B1A1-438E-A425-CEBC7DC684EE}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}" -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Shell Extensions\Approved\" "{E768CD3B-BDDC-436D-9C13-E1B39CA257B1}" -Force -ea 0
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{31D09BA0-12F5-4CCE-BE8A-2923E76605DA}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{B4F3A835-0E21-4959-BA22-42B3008E02FF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Explorer\Browser Helper Objects\{D0498E0A-45B7-42AE-A9AA-ABA463DBD3BF}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\NameSpace\{0875DCB6-C686-4243-9432-ADCCF0B9F2D7}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Desktop\Namespace\{B28AA736-876B-46DA-B3A8-84C5E30BA492}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\NetworkNeighborhood\Namespace\{46137B78-0EC3-426D-8B89-FF7C3A458B5E}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\VolumeCaches\Microsoft Office Temp Files" -Force -ea 0 -Recurse
}
function Get-MsiProducts {
  
  # PowerShell: Get-MsiProducts / Get Windows Installer Products
  # https://gist.github.com/MyITGuy/153fc0f553d840631269720a56be5136#file-file01-ps1

    function Get-MsiUpgradeCode {
        [CmdletBinding()]
        param (
            [System.Guid]$ProductCode
            ,
            [System.Guid]$UpgradeCode
        )
        function ConvertFrom-CompressedGuid {
            <#
	        .SYNOPSIS
		        Converts a compressed globally unique identifier (GUID) string into a GUID string.
	        .DESCRIPTION
            Takes a compressed GUID string and breaks it into 6 parts. It then loops through the first five parts and reversing the order. It loops through the sixth part and reversing the order of every 2 characters. It then joins the parts back together and returns a GUID.
	        .EXAMPLE
		        ConvertFrom-CompressedGuid -CompressedGuid '2820F6C7DCD308A459CABB92E828C144'
	
		        The output of this example would be: {7C6F0282-3DCD-4A80-95AC-BB298E821C44}
	        .PARAMETER CompressedGuid
		        A compressed globally unique identifier (GUID) string.
	        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
                [ValidatePattern('^[0-9a-fA-F]{32}$')]
                [ValidateScript( { [System.Guid]::Parse($_) -is [System.Guid] })]
                [System.String]$CompressedGuid
            )
            process {
                Write-Verbose "CompressedGuid: $($CompressedGuid)"
                $GuidString = ([System.Guid]$CompressedGuid).ToString('N')
                Write-Verbose "GuidString: $($GuidString)"
                $Indexes = [ordered]@{
                    0  = 8
                    8  = 4
                    12 = 4
                    16 = 2
                    18 = 2
                    20 = 12
                }
                $Guid = ''
                foreach ($key in $Indexes.Keys) {
                    $value = $Indexes[$key]
                    $Substring = $GuidString.Substring($key, $value)
                    Write-Verbose "Substring: $($Substring)"
                    switch ($key) {
                        20 {
                            $parts = $Substring -split '(.{2})' | Where-Object { $_ }
                            foreach ($part In $parts) {
                                $part = $part -split '(.{1})'
                                Write-Verbose "Part: $($part)"
                                [System.Array]::Reverse($part)
                                Write-Verbose "Reversed: $($part)"
                                $Guid += $part -join ''
                            }
                        }
                        default {
                            $part = $Substring.ToCharArray()
                            Write-Verbose "Part: $($part)"
                            [System.Array]::Reverse($part)
                            Write-Verbose "Reversed: $($part)"
                            $Guid += $part -join ''
                        }
                    }
                }
                [System.Guid]::Parse($Guid).ToString('B').ToUpper()
            }
        }

        function ConvertTo-CompressedGuid {
            <#
	        .SYNOPSIS
		        Converts a GUID string into a compressed globally unique identifier (GUID) string.
	        .DESCRIPTION
		        Takes a GUID string and breaks it into 6 parts. It then loops through the first five parts and reversing the order. It loops through the sixth part and reversing the order of every 2 characters. It then joins the parts back together and returns a compressed GUID string.
	        .EXAMPLE
		        ConvertTo-CompressedGuid -Guid '{7C6F0282-3DCD-4A80-95AC-BB298E821C44}'
	
            The output of this example would be: 2820F6C7DCD308A459CABB92E828C144
	        .PARAMETER Guid
            A globally unique identifier (GUID).
	        #>
            [CmdletBinding()]
            [OutputType([System.String])]
            param (
                [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
                [ValidateScript( { [System.Guid]::Parse($_) -is [System.Guid] })]
                [System.Guid]$Guid
            )
            process {
                Write-Verbose "Guid: $($Guid)"
                $GuidString = $Guid.ToString('N')
                Write-Verbose "GuidString: $($GuidString)"
                $Indexes = [ordered]@{
                    0  = 8
                    8  = 4
                    12 = 4
                    16 = 2
                    18 = 2
                    20 = 12
                }
                $CompressedGuid = ''
                foreach ($key in $Indexes.keys) {
                    $value = $Indexes[$key]
                    $Substring = $GuidString.Substring($key, $value)
                    Write-Verbose "Substring: $($Substring)"
                    switch ($key) {
                        20 {
                            $parts = $Substring -split '(.{2})' | Where-Object { $_ }
                            foreach ($part In $parts) {
                                $part = $part -split '(.{1})'
                                Write-Verbose "Part: $($part)"
                                [System.Array]::Reverse($part)
                                Write-Verbose "Reversed: $($part)"
                                $CompressedGuid += $part -join ''
                            }
                        }
                        default {
                            $part = $Substring.ToCharArray()
                            Write-Verbose "Part: $($part)"
                            [System.Array]::Reverse($part)
                            Write-Verbose "Reversed: $($part)"
                            $CompressedGuid += $part -join ''
                        }
                    }
                }
                [System.Guid]::Parse($CompressedGuid).ToString('N').ToUpper()
            }
        }

        filter ByProductCode {
            $Object = $_
            Write-Verbose "ProductCode: $($ProductCode)"
            if ($ProductCode) {
                $Object | Where-Object { [System.Guid]($_.ProductCode) -eq [System.Guid]($ProductCode) }
                break
            }
            $Object
        }

        $Path = "Registry::HKEY_CLASSES_ROOT\Installer\UpgradeCodes\*"
        if ($UpgradeCode) {
            $CompressedUpgradeCode = ConvertTo-CompressedGuid -Guid $UpgradeCode -Verbose:$false
            Write-Verbose "CompressedUpgradeCode: $($CompressedUpgradeCode)"
            $Path = "Registry::HKEY_CLASSES_ROOT\Installer\UpgradeCodes\$($CompressedUpgradeCode)"
        }

        Get-Item -Path $Path -ErrorAction SilentlyContinue | ForEach-Object {
            $UpgradeCodeFromCompressedGuid = ConvertFrom-CompressedGuid -CompressedGuid $_.PSChildName -Verbose:$false
            foreach ($ProductCodeCompressedGuid in ($_.GetValueNames())) {
                $Properties = [ordered]@{
                    ProductCode = ConvertFrom-CompressedGuid -CompressedGuid $ProductCodeCompressedGuid -Verbose:$false
                    UpgradeCode = [System.Guid]::Parse($UpgradeCodeFromCompressedGuid).ToString('B').ToUpper()
                }
                [PSCustomObject]$Properties | ByProductCode
            }
        }
    }

    $MsiUpgradeCodes = Get-MsiUpgradeCode

    $Installer = New-Object -ComObject WindowsInstaller.Installer
	$Type = $Installer.GetType()
	$Products = $Type.InvokeMember('Products', [System.Reflection.BindingFlags]::GetProperty, $null, $Installer, $null)
	foreach ($Product In $Products) {
		$hash = @{}
		$hash.ProductCode = $Product
		$Attributes = @('Language', 'ProductName', 'PackageCode', 'Transforms', 'AssignmentType', 'PackageName', 'InstalledProductName', 'VersionString', 'RegCompany', 'RegOwner', 'ProductID', 'ProductIcon', 'InstallLocation', 'InstallSource', 'InstallDate', 'Publisher', 'LocalPackage', 'HelpLink', 'HelpTelephone', 'URLInfoAbout', 'URLUpdateInfo')		
		foreach ($Attribute In $Attributes) {
			$hash."$($Attribute)" = $null
		}
		foreach ($Attribute In $Attributes) {
			try {
				$hash."$($Attribute)" = $Type.InvokeMember('ProductInfo', [System.Reflection.BindingFlags]::GetProperty, $null, $Installer, @($Product, $Attribute))
			} catch [System.Exception] {
				#$error[0]|format-list –force
			}
		}
        
        # UpgradeCode
        $hash.UpgradeCode = $MsiUpgradeCodes | Where-Object ProductCode -eq ($hash.ProductCode) | Select-Object -ExpandProperty UpgradeCode

		New-Object -TypeName PSObject -Property $hash
	}
}
function UninstallLicenses($DllPath) {
  
  # https://github.com/ave9858
  # https://gist.github.com/ave9858/9fff6af726ba3ddc646285d1bbf37e71

    $DynAssembly = New-Object System.Reflection.AssemblyName('Win32Lib')
    $AssemblyBuilder = [AppDomain]::CurrentDomain.DefineDynamicAssembly($DynAssembly, [Reflection.Emit.AssemblyBuilderAccess]::Run)
    $ModuleBuilder = $AssemblyBuilder.DefineDynamicModule('Win32Lib', $False)
    $TypeBuilder = $ModuleBuilder.DefineType('sppc', 'Public, Class')
    $DllImportConstructor = [Runtime.InteropServices.DllImportAttribute].GetConstructor(@([String]))
    $FieldArray = [Reflection.FieldInfo[]] @([Runtime.InteropServices.DllImportAttribute].GetField('EntryPoint'))

    $Open = $TypeBuilder.DefineMethod('SLOpen', [Reflection.MethodAttributes] 'Public, Static', [int], @([IntPtr].MakeByRefType()))
    $Open.SetCustomAttribute((New-Object Reflection.Emit.CustomAttributeBuilder(
                $DllImportConstructor,
                @($DllPath),
                $FieldArray,
                @('SLOpen'))))

    $GetSLIDList = $TypeBuilder.DefineMethod('SLGetSLIDList', [Reflection.MethodAttributes] 'Public, Static', [int], @([IntPtr], [int], [guid].MakeByRefType(), [int], [int].MakeByRefType(), [IntPtr].MakeByRefType()))
    $GetSLIDList.SetCustomAttribute((New-Object Reflection.Emit.CustomAttributeBuilder(
                $DllImportConstructor,
                @($DllPath),
                $FieldArray,
                @('SLGetSLIDList'))))

    $UninstallLicense = $TypeBuilder.DefineMethod('SLUninstallLicense', [Reflection.MethodAttributes] 'Public, Static', [int], @([IntPtr], [IntPtr]))
    $UninstallLicense.SetCustomAttribute((New-Object Reflection.Emit.CustomAttributeBuilder(
                $DllImportConstructor,
                @($DllPath),
                $FieldArray,
                @('SLUninstallLicense'))))

    $SPPC = $TypeBuilder.CreateType()
    $Handle = [IntPtr]::Zero
    $SPPC::SLOpen([ref]$handle) | Out-Null
    $pnReturnIds = 0
    $ppReturnIds = [IntPtr]::Zero

    if (!$SPPC::SLGetSLIDList($handle, 0, [ref][guid]"0ff1ce15-a989-479d-af46-f275c6370663", 6, [ref]$pnReturnIds, [ref]$ppReturnIds)) {
        foreach ($i in 0..($pnReturnIds - 1)) {
            $SPPC::SLUninstallLicense($handle, [System.Int64]$ppReturnIds + [System.Int64]16 * $i) | Out-Null
        }    
    }
}
function GetUninstall {
$UninstallArr  = @{}
$UninstallKeys = @{}
$UninstallKeys.Add(1,"HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
$UninstallKeys.Add(2,"HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")

foreach ($sKey in $UninstallKeys.Values) {
  Set-Location "HKLM:\"
  Push-Location $sKey -ea 0

  if (@(Get-Location).Path -NE 'HKLM:\') {
    $children = gci .
    $children | % {
      $sName = $_.Name.Replace('HKEY_LOCAL_MACHINE','HKLM:')
      $sGuid = $sName.Split('\')|select -Last 1
      Set-Location "HKLM:\"; Push-Location "$sName" -ea 0
      if (@(Get-Location).Path -NE 'HKLM:\') {
        try {
          $UninstallString = $null
          $UninstallString = gpv . -Name 'UninstallString' -ea 0 }
        catch {}
        if ($UninstallString -and (
          $UninstallString|IsC2R)) {

            try {
              $UninstallArr.Add(
                $sGuid, $UninstallString)}
            catch {}}}}}
}

return $UninstallArr
}
function CleanUninstall {

$UninstallKeys = @{}
$UninstallKeys.Add(1,"HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
$UninstallKeys.Add(2,"HKLM:SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall")

foreach ($sKey in $UninstallKeys.Values) {
Set-Location "HKLM:\"
Push-Location $sKey -ea 0

if (@(Get-Location).Path -NE 'HKLM:\') {
  $children = gci .
  $children | % {
    $sName = $_.Name.Replace('HKEY_LOCAL_MACHINE','HKLM:')
    $sGuid = $sName.Split('\')|select -Last 1
    
    Set-Location "HKLM:\"
    Push-Location "$sName" -ea 0
    if (@(Get-Location).Path -NE 'HKLM:\') {
      try {
        $InstallLocation = $null
        $InstallLocation = gpv . -Name 'InstallLocation' -ea 0 }
      catch {}

      if (($sGuid -and ($sGuid|CheckDelete)) -or (
        $InstallLocation -and ($InstallLocation|IsC2R))) {
          Set-Location "HKLM:\"
          RI $sName -Recurse -Force }
    }}}}
}
Function RegWipe {

CloseOfficeApps

"*** -- C2R specifics"
"*** -- Virtual InstallRoot"
"*** -- Mapi Search reg"
"*** -- Office key in HKLM"

Set-Location "HKLM:\"
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Run" Lync15 -Force -ea 0
RP "HKLM:SOFTWARE\Microsoft\Windows\CurrentVersion\Run" Lync16 -Force -ea 0
RI "HKLM:SOFTWARE\Microsoft\Office\15.0\Common\InstallRoot\Virtual" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\16.0\Common\InstallRoot\Virtual" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\Common\InstallRoot\Virtual" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Classes\CLSID\{2027FC3B-CF9D-4ec7-A823-38BA308625CC}" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\15.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\15.0\ClickToRunStore" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\16.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\16.0\ClickToRunStore" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\ClickToRunStore" -Force -ea 0 -Recurse
RI "HKLM:Software\Microsoft\Office\15.0" -Force -ea 0 -Recurse
RI "HKLM:Software\Microsoft\Office\16.0" -Force -ea 0 -Recurse

"*** -- HKCU Registration"
Set-Location "HKCU:\"
RI "HKCU:Software\Microsoft\Office\15.0\Registration" -Force -ea 0 -Recurse
RI "HKCU:Software\Microsoft\Office\16.0\Registration" -Force -ea 0 -Recurse
RI "HKCU:Software\Microsoft\Office\Registration" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\15.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\16.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:Software\Microsoft\Office\15.0" -Force -ea 0 -Recurse
RI "HKCU:Software\Microsoft\Office\16.0" -Force -ea 0 -Recurse

"*** -- App Paths"
$Keys = reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths" 2>$null
$keys | % {
$value = reg query "$_" /ve /t REG_SZ 2>$null
if ($value -match "\\Microsoft Office") {
  reg delete $_ /f | Out-Null }}

"*** -- Run key"
$hDefKey = "HKLM"
$sSubKeyName = "SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
Set-Location "$($hDefKey):\"
Push-Location "$($hDefKey):$($sSubKeyName)" -ea 0

if (@(Get-Location).Path -ne "$($hDefKey):\") {
  $arrNames = gi .
  if ($arrNames)  {
    $arrNames.Property | % { 
      $name = GPV . $_
      if ($name -and (
        $Name|IsC2R)) {
          RP . $_ -Force
}}}}

"*** -- Un-install Keys"
CleanUninstall

"*** -- UpgradeCodes, WI config, WI global config"
"*** -- msiexec based uninstall [Fail-Safe]"

# First here ... 
# HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products

$hash     = $null;
$HashList = $null;
$sKey     = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products"

Set-Location 'HKLM:\'
$sProducts = 
  GCI $sKey -ea 0
$HashList = $sProducts | % {
  ($_).PSPath.Split('\') | select -Last 1 | % {
    [PSCustomObject]@{
    cGuid = $_
    sGuid = ($_|GetExpandedGuid) }}}
$GuidList = 
  $HashList | ? sGuid

if ($GuidList) {
  $GuidList | ? sGuid | % {
    $Proc = $null
    $ProductCode = $_.sGuid
    $sMsiProp = "REBOOT=ReallySuppress NOREMOVESPAWN=True"
    $sUninstallCmd = "/x {$($ProductCode)} $($sMsiProp) /q"

    if ($ProductCode) {
      $Proc = start msiexec.exe -Args $sUninstallCmd -Wait -WindowStyle Hidden -ea 0 -PassThru
      "*** -- Msiexec $($sUninstallCmd) ,End with value: $($proc.ExitCode)" }

    Set-Location 'HKLM:\'
    RI "$sKey\$($_.sGuid)" -Force -Recurse -ea 0 | Out-Null
    Set-Location 'HKCR:\'
    RI "HKCR:\Installer\Products\$($_.sGuid)" -Force -Recurse -ea 0 | Out-Null }}

# Second here ... 
# HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes

$hash     = $null;
$HashList = $null;
$sKey     = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes"

Set-Location 'HKLM:\'
$sUpgradeCodes = 
  GCI $sKey -ea 0
$HashList = $sUpgradeCodes | % {
  ($_).PSPath.Split('\') | select -Last 1 | % {
    [PSCustomObject]@{
    cGuid = $_
    sGuid = ($_|GetExpandedGuid) }}}
$GuidList = 
  $HashList | ? sGuid

if ($GuidList) {
  $GuidList | % {
    Set-Location 'HKLM:\'
    RI "$sKey\$($_.sGuid)" -Force -Recurse -ea 0 | Out-Null
    Set-Location 'HKCR:\'
    RI "HKCR:\Installer\UpgradeCodes\$($_.sGuid)" -Force -Recurse -ea 0 | Out-Null }}

# make sure we clean everything
$sKeyToRe = @{}
$sKeyList = (
  "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UpgradeCodes",
  "SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products" )

foreach ($sKey in $sKeyList)
{
  Set-Location "HKLM:\"
  $sKey = "HKLM:" + $sKey

  Set-Location "HKLM:\"
  Push-Location $sKey -ea 0

  if (@(Get-Location).Path -NE 'HKLM:\') {
    $children = gci .
    $children | % {
      $sName = $_.Name.Replace('HKEY_LOCAL_MACHINE','HKLM:')
      $sGuid = $sName.Split('\')|select -Last 1

      Set-Location "HKLM:\"
      Push-Location $sName -ea 0
      if (@(Get-Location).Path -NE 'HKLM:\') {

      $InstallSource   = $null
      $UninstallString = $null
    
      try {
        $InstallSource   = GPV "InstallProperties" -Name InstallSource   -ea 0
        $UninstallString = GPV "InstallProperties" -Name UninstallString -ea 0 }
      catch { }
    
      $CheckOfficeApp = $null
      $CheckOfficeApp = ($sGuid -and ($sGuid|CheckDelete)) -or (
        $InstallSource -and $UninstallString -and ($InstallSource|ISC2R) -and (
        [REGEX]::Match($UninstallString, "^.*{(.*)}.*$",$IgnoreCase)))

      if ($CheckOfficeApp -eq $true) {
        $Matches = [REGEX]::Matches($UninstallString,"^.*{(.*)}.*$",
          $IgnoreCase)
        try {
          $ProductCode = $null
          $ProductCode = $Matches[0].Groups[1].Value }
        catch {}

        $proc = $null
        $sMsiProp = "REBOOT=ReallySuppress NOREMOVESPAWN=True"
        $sUninstallCmd = "/x {$($ProductCode)} $($sMsiProp) /q"

        if ($ProductCode) {
          $proc = start msiexec.exe -Args $sUninstallCmd -Wait -WindowStyle Hidden -ea 0 -PassThru
          "*** -- mSiexec $($sUninstallCmd) ,End with value: $($proc.ExitCode)"
		  $sKeyToRe.Add($sName,$sName) }
}}}}}

Set-Location "HKLM:\"
$sKeyToRe.Keys | % {
  RI $_ -Force -Recurse -ea 0 | Out-Null }

Set-Location "HKCR:\"
$sKeyToRe.Keys | % {
  $GUID = ($_).Split('\') | Select-Object -Last 1
  if ($GUID) {
    RI "HKCR:\Installer\Products\$GUID" -Force -Recurse -ea 0 | Out-Null }}

"*** -- Known Typelib Registration"
RegWipeTypeLib

"*** -- Published Components [JAWOT]"

"*** -- ActiveX/COM Components [JAWOT]"
$COM = (
"{00020800-0000-0000-C000-000000000046}","{00020803-0000-0000-C000-000000000046}",
"{00020812-0000-0000-C000-000000000046}","{00020820-0000-0000-C000-000000000046}",
"{00020821-0000-0000-C000-000000000046}","{00020827-0000-0000-C000-000000000046}",
"{00020830-0000-0000-C000-000000000046}","{00020832-0000-0000-C000-000000000046}",
"{00020833-0000-0000-C000-000000000046}","{00020906-0000-0000-C000-000000000046}",
"{00020907-0000-0000-C000-000000000046}","{000209F0-0000-0000-C000-000000000046}",
"{000209F4-0000-0000-C000-000000000046}","{000209F5-0000-0000-C000-000000000046}",
"{000209FE-0000-0000-C000-000000000046}","{000209FF-0000-0000-C000-000000000046}",
"{00024500-0000-0000-C000-000000000046}","{00024502-0000-0000-C000-000000000046}",
"{00024505-0016-0000-C000-000000000046}","{048EB43E-2059-422F-95E0-557DA96038AF}",
"{18A06B6B-2F3F-4E2B-A611-52BE631B2D22}","{1B261B22-AC6A-4E68-A870-AB5080E8687B}",
"{1CDC7D25-5AA3-4DC4-8E0C-91524280F806}","{3C18EAE4-BC25-4134-B7DF-1ECA1337DDDC}",
"{64818D10-4F9B-11CF-86EA-00AA00B929E8}","{64818D11-4F9B-11CF-86EA-00AA00B929E8}",
"{65235197-874B-4A07-BDC5-E65EA825B718}","{73720013-33A0-11E4-9B9A-00155D152105}",
"{75D01070-1234-44E9-82F6-DB5B39A47C13}","{767A19A0-3CC7-415B-9D08-D48DD7B8557D}",
"{84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}","{8A624388-AA27-43E0-89F8-2A12BFF7BCCD}",
"{912ABC52-36E2-4714-8E62-A8B73CA5E390}","{91493441-5A91-11CF-8700-00AA0060263B}",
"{AA14F9C9-62B5-4637-8AC4-8F25BF29D5A7}","{C282417B-2662-44B8-8A94-3BFF61C50900}",
"{CF4F55F4-8F87-4D47-80BB-5808164BB3F8}","{DC020317-E6E2-4A62-B9FA-B3EFE16626F4}",
"{EABCECDB-CC1C-4A6F-B4E3-7F888A5ADFC8}","{F4754C9B-64F5-4B40-8AF4-679732AC0607}")

#Set-Location "HKCR:\"
$COM | % {
  # will not work .. why ? don't know
  # ri "HKCR:CLSID\$_" -Recurse -Force -ea 0 
}

"*** -- TypeLib Interface [JAWOT]"
$interface = @(
"{000672AC-0000-0000-C000-000000000046}","{000C0300-0000-0000-C000-000000000046}"
"{000C0301-0000-0000-C000-000000000046}","{000C0302-0000-0000-C000-000000000046}"
"{000C0304-0000-0000-C000-000000000046}","{000C0306-0000-0000-C000-000000000046}"
"{000C0308-0000-0000-C000-000000000046}","{000C030A-0000-0000-C000-000000000046}"
"{000C030C-0000-0000-C000-000000000046}","{000C030D-0000-0000-C000-000000000046}"
"{000C030E-0000-0000-C000-000000000046}","{000C0310-0000-0000-C000-000000000046}"
"{000C0311-0000-0000-C000-000000000046}","{000C0312-0000-0000-C000-000000000046}"
"{000C0313-0000-0000-C000-000000000046}","{000C0314-0000-0000-C000-000000000046}"
"{000C0315-0000-0000-C000-000000000046}","{000C0316-0000-0000-C000-000000000046}"
"{000C0317-0000-0000-C000-000000000046}","{000C0318-0000-0000-C000-000000000046}"
"{000C0319-0000-0000-C000-000000000046}","{000C031A-0000-0000-C000-000000000046}"
"{000C031B-0000-0000-C000-000000000046}","{000C031C-0000-0000-C000-000000000046}"
"{000C031D-0000-0000-C000-000000000046}","{000C031E-0000-0000-C000-000000000046}"
"{000C031F-0000-0000-C000-000000000046}","{000C0320-0000-0000-C000-000000000046}"
"{000C0321-0000-0000-C000-000000000046}","{000C0322-0000-0000-C000-000000000046}"
"{000C0324-0000-0000-C000-000000000046}","{000C0326-0000-0000-C000-000000000046}"
"{000C0328-0000-0000-C000-000000000046}","{000C032E-0000-0000-C000-000000000046}"
"{000C0330-0000-0000-C000-000000000046}","{000C0331-0000-0000-C000-000000000046}"
"{000C0332-0000-0000-C000-000000000046}","{000C0333-0000-0000-C000-000000000046}"
"{000C0334-0000-0000-C000-000000000046}","{000C0337-0000-0000-C000-000000000046}"
"{000C0338-0000-0000-C000-000000000046}","{000C0339-0000-0000-C000-000000000046}"
"{000C033A-0000-0000-C000-000000000046}","{000C033B-0000-0000-C000-000000000046}"
"{000C033D-0000-0000-C000-000000000046}","{000C033E-0000-0000-C000-000000000046}"
"{000C0340-0000-0000-C000-000000000046}","{000C0341-0000-0000-C000-000000000046}"
"{000C0353-0000-0000-C000-000000000046}","{000C0356-0000-0000-C000-000000000046}"
"{000C0357-0000-0000-C000-000000000046}","{000C0358-0000-0000-C000-000000000046}"
"{000C0359-0000-0000-C000-000000000046}","{000C035A-0000-0000-C000-000000000046}"
"{000C0360-0000-0000-C000-000000000046}","{000C0361-0000-0000-C000-000000000046}"
"{000C0362-0000-0000-C000-000000000046}","{000C0363-0000-0000-C000-000000000046}"
"{000C0364-0000-0000-C000-000000000046}","{000C0365-0000-0000-C000-000000000046}"
"{000C0366-0000-0000-C000-000000000046}","{000C0367-0000-0000-C000-000000000046}"
"{000C0368-0000-0000-C000-000000000046}","{000C0369-0000-0000-C000-000000000046}"
"{000C036A-0000-0000-C000-000000000046}","{000C036C-0000-0000-C000-000000000046}"
"{000C036D-0000-0000-C000-000000000046}","{000C036E-0000-0000-C000-000000000046}"
"{000C036F-0000-0000-C000-000000000046}","{000C0370-0000-0000-C000-000000000046}"
"{000C0371-0000-0000-C000-000000000046}","{000C0372-0000-0000-C000-000000000046}"
"{000C0373-0000-0000-C000-000000000046}","{000C0375-0000-0000-C000-000000000046}"
"{000C0376-0000-0000-C000-000000000046}","{000C0377-0000-0000-C000-000000000046}"
"{000C0379-0000-0000-C000-000000000046}","{000C037A-0000-0000-C000-000000000046}"
"{000C037B-0000-0000-C000-000000000046}","{000C037C-0000-0000-C000-000000000046}"
"{000C037D-0000-0000-C000-000000000046}","{000C037E-0000-0000-C000-000000000046}"
"{000C037F-0000-0000-C000-000000000046}","{000C0380-0000-0000-C000-000000000046}"
"{000C0381-0000-0000-C000-000000000046}","{000C0382-0000-0000-C000-000000000046}"
"{000C0385-0000-0000-C000-000000000046}","{000C0386-0000-0000-C000-000000000046}"
"{000C0387-0000-0000-C000-000000000046}","{000C0388-0000-0000-C000-000000000046}"
"{000C0389-0000-0000-C000-000000000046}","{000C038A-0000-0000-C000-000000000046}"
"{000C038B-0000-0000-C000-000000000046}","{000C038C-0000-0000-C000-000000000046}"
"{000C038E-0000-0000-C000-000000000046}","{000C038F-0000-0000-C000-000000000046}"
"{000C0390-0000-0000-C000-000000000046}","{000C0391-0000-0000-C000-000000000046}"
"{000C0392-0000-0000-C000-000000000046}","{000C0393-0000-0000-C000-000000000046}"
"{000C0395-0000-0000-C000-000000000046}","{000C0396-0000-0000-C000-000000000046}"
"{000C0397-0000-0000-C000-000000000046}","{000C0398-0000-0000-C000-000000000046}"
"{000C0399-0000-0000-C000-000000000046}","{000C039A-0000-0000-C000-000000000046}"
"{000C03A0-0000-0000-C000-000000000046}","{000C03A1-0000-0000-C000-000000000046}"
"{000C03A2-0000-0000-C000-000000000046}","{000C03A3-0000-0000-C000-000000000046}"
"{000C03A4-0000-0000-C000-000000000046}","{000C03A5-0000-0000-C000-000000000046}"
"{000C03A6-0000-0000-C000-000000000046}","{000C03A7-0000-0000-C000-000000000046}"
"{000C03B2-0000-0000-C000-000000000046}","{000C03B9-0000-0000-C000-000000000046}"
"{000C03BA-0000-0000-C000-000000000046}","{000C03BB-0000-0000-C000-000000000046}"
"{000C03BC-0000-0000-C000-000000000046}","{000C03BD-0000-0000-C000-000000000046}"
"{000C03BE-0000-0000-C000-000000000046}","{000C03BF-0000-0000-C000-000000000046}"
"{000C03C0-0000-0000-C000-000000000046}","{000C03C1-0000-0000-C000-000000000046}"
"{000C03C2-0000-0000-C000-000000000046}","{000C03C3-0000-0000-C000-000000000046}"
"{000C03C4-0000-0000-C000-000000000046}","{000C03C5-0000-0000-C000-000000000046}"
"{000C03C6-0000-0000-C000-000000000046}","{000C03C7-0000-0000-C000-000000000046}"
"{000C03C8-0000-0000-C000-000000000046}","{000C03C9-0000-0000-C000-000000000046}"
"{000C03CA-0000-0000-C000-000000000046}","{000C03CB-0000-0000-C000-000000000046}"
"{000C03CC-0000-0000-C000-000000000046}","{000C03CD-0000-0000-C000-000000000046}"
"{000C03CE-0000-0000-C000-000000000046}","{000C03CF-0000-0000-C000-000000000046}"
"{000C03D0-0000-0000-C000-000000000046}","{000C03D1-0000-0000-C000-000000000046}"
"{000C03D2-0000-0000-C000-000000000046}","{000C03D3-0000-0000-C000-000000000046}"
"{000C03D4-0000-0000-C000-000000000046}","{000C03D5-0000-0000-C000-000000000046}"
"{000C03D6-0000-0000-C000-000000000046}","{000C03D7-0000-0000-C000-000000000046}"
"{000C03E0-0000-0000-C000-000000000046}","{000C03E1-0000-0000-C000-000000000046}"
"{000C03E2-0000-0000-C000-000000000046}","{000C03E3-0000-0000-C000-000000000046}"
"{000C03E4-0000-0000-C000-000000000046}","{000C03E5-0000-0000-C000-000000000046}"
"{000C03E6-0000-0000-C000-000000000046}","{000C03F0-0000-0000-C000-000000000046}"
"{000C03F1-0000-0000-C000-000000000046}","{000C0410-0000-0000-C000-000000000046}"
"{000C0411-0000-0000-C000-000000000046}","{000C0913-0000-0000-C000-000000000046}"
"{000C0914-0000-0000-C000-000000000046}","{000C0936-0000-0000-C000-000000000046}"
"{000C1530-0000-0000-C000-000000000046}","{000C1531-0000-0000-C000-000000000046}"
"{000C1532-0000-0000-C000-000000000046}","{000C1533-0000-0000-C000-000000000046}"
"{000C1534-0000-0000-C000-000000000046}","{000C1709-0000-0000-C000-000000000046}"
"{000C170B-0000-0000-C000-000000000046}","{000C170F-0000-0000-C000-000000000046}"
"{000C1710-0000-0000-C000-000000000046}","{000C1711-0000-0000-C000-000000000046}"
"{000C1712-0000-0000-C000-000000000046}","{000C1713-0000-0000-C000-000000000046}"
"{000C1714-0000-0000-C000-000000000046}","{000C1715-0000-0000-C000-000000000046}"
"{000C1716-0000-0000-C000-000000000046}","{000C1717-0000-0000-C000-000000000046}"
"{000C1718-0000-0000-C000-000000000046}","{000C171B-0000-0000-C000-000000000046}"
"{000C171C-0000-0000-C000-000000000046}","{000C1723-0000-0000-C000-000000000046}"
"{000C1724-0000-0000-C000-000000000046}","{000C1725-0000-0000-C000-000000000046}"
"{000C1726-0000-0000-C000-000000000046}","{000C1727-0000-0000-C000-000000000046}"
"{000C1728-0000-0000-C000-000000000046}","{000C1729-0000-0000-C000-000000000046}"
"{000C172A-0000-0000-C000-000000000046}","{000C172B-0000-0000-C000-000000000046}"
"{000C172C-0000-0000-C000-000000000046}","{000C172D-0000-0000-C000-000000000046}"
"{000C172E-0000-0000-C000-000000000046}","{000C172F-0000-0000-C000-000000000046}"
"{000C1730-0000-0000-C000-000000000046}","{000C1731-0000-0000-C000-000000000046}"
"{000CD100-0000-0000-C000-000000000046}","{000CD101-0000-0000-C000-000000000046}"
"{000CD102-0000-0000-C000-000000000046}","{000CD6A1-0000-0000-C000-000000000046}"
"{000CD6A2-0000-0000-C000-000000000046}","{000CD6A3-0000-0000-C000-000000000046}"
"{000CD706-0000-0000-C000-000000000046}","{000CD809-0000-0000-C000-000000000046}"
"{000CD900-0000-0000-C000-000000000046}","{000CD901-0000-0000-C000-000000000046}"
"{000CD902-0000-0000-C000-000000000046}","{000CD903-0000-0000-C000-000000000046}"
"{000CDB00-0000-0000-C000-000000000046}","{000CDB01-0000-0000-C000-000000000046}"
"{000CDB02-0000-0000-C000-000000000046}","{000CDB03-0000-0000-C000-000000000046}"
"{000CDB04-0000-0000-C000-000000000046}","{000CDB05-0000-0000-C000-000000000046}"
"{000CDB06-0000-0000-C000-000000000046}","{000CDB09-0000-0000-C000-000000000046}"
"{000CDB0A-0000-0000-C000-000000000046}","{000CDB0E-0000-0000-C000-000000000046}"
"{000CDB0F-0000-0000-C000-000000000046}","{000CDB10-0000-0000-C000-000000000046}"
"{00194002-D9C3-11D3-8D59-0050048384E3}","{4291224C-DEFE-485B-8E69-6CF8AA85CB76}"
"{4B0F95AC-5769-40E9-98DF-80CDD086612E}","{4CAC6328-B9B0-11D3-8D59-0050048384E3}"
"{55F88890-7708-11D1-ACEB-006008961DA5}","{55F88892-7708-11D1-ACEB-006008961DA5}"
"{55F88896-7708-11D1-ACEB-006008961DA5}","{6EA00553-9439-4D5A-B1E6-DC15A54DA8B2}"
"{88FF5F69-FACF-4667-8DC8-A85B8225DF15}","{8A64A872-FC6B-4D4A-926E-3A3689562C1C}"
"{919AA22C-B9AD-11D3-8D59-0050048384E3}","{A98639A1-CB0C-4A5C-A511-96547F752ACD}"
"{ABFA087C-F703-4D53-946E-37FF82B2C994}","{D996597A-0E80-4753-81FC-DCF16BDF4947}"
"{DE9CD4FF-754A-49DD-A0DC-B787DA2DB0A1}","{DFD3BED7-93EC-4BCE-866C-6BAB41D28621}"
)

#Set-Location "HKCR:\"
$interface | % {
  # will not work .. why ? don't know
  # RI "HKCR\Interface\$_" -Recurse -Force -ea 0
}

"*** -- Components in Global [ & Could take 2-3 minutes & ]"
$Keys = reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Components" 2>$null
$keys | % {
 $data = reg query $_ /t REG_SZ 2>$null
 if (($data -ne $null) -and (
   $data -match "\\Microsoft Office")) {
     reg delete $_ /f | Out-Null }}

"*** -- Components in CLSID [ & Could take 2-3 minutes & ]"
$Keys = reg query "HKLM\SOFTWARE\Classes\CLSID" 2>$null
$keys | % {
  $LocalServer32 = reg query "$_\LocalServer32" /ve /t REG_SZ 2>$null
  if (($LocalServer32 -ne $null) -and (
    $LocalServer32[2] -match "\\Microsoft Office")) {
      reg delete $_ /f | Out-Null }
  if ($LocalServer32 -eq $null) {
    $InprocServer32 = reg query "$_\InprocServer32" /ve /t REG_SZ 2>$null
    if (($InprocServer32 -ne $null) -and (
      $InprocServer32[2] -match "\\Microsoft Office")) {
        reg delete $_ /f | Out-Null }}}

<#
-- reg query "HKEY_CLASSES_ROOT\CLSID\
-- "HKEY_CLASSES_ROOT\CLSID\{C282417B-2662-44B8-8A94-3BFF61C50900}"

-- reg query "HKEY_CLASSES_ROOT\CLSID\{C282417B-2662-44B8-8A94-3BFF61C50900}\LocalServer32"
-- ERROR: The system was unable to find the specified registry key or value. [ACCESS DENIED ERROR]

$Keys = reg query "HKCR\CLSID" 2>$null
$keys | % {
  $LocalServer32 = reg query "$_\LocalServer32" /ve /t REG_SZ 2>$null
  if (($LocalServer32 -ne $null) -and (
    $LocalServer32[2] -match "\\Microsoft Office")) {
      reg delete $_ /f | Out-Null }
  if ($LocalServer32 -eq $null) {
    $InprocServer32 = reg query "$_\InprocServer32" /ve /t REG_SZ 2>$null
    if (($InprocServer32 -ne $null) -and (
      $InprocServer32[2] -match "\\Microsoft Office")) {
        reg delete $_ /f | Out-Null }}}
#>
}
Function FileWipe {

"*** -- remove the OfficeSvc service"
$service = $null
$service = Get-WmiObject Win32_Service -Filter "Name='OfficeSvc'" -ea 0
if ($service) { 
  try {
    $service.delete()|out-null}
  catch {} }

"*** -- remove the ClickToRunSvc service"
$service = $null
$service = Get-WmiObject Win32_Service -Filter "Name='ClickToRunSvc'" -ea 0
if ($service) { 
  try {
    $service.delete()|out-null}
  catch {} }

"*** -- delete C2R package files"
Set-Location "$($env:SystemDrive)\"

RI @(Join-Path $env:ProgramFiles "Microsoft Office\Office16") -Recurse -force -ea 0
RI @(Join-Path $env:ProgramData "Microsoft\office\FFPackageLocker") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramData "Microsoft\office\FFStatePBLocker") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\AppXManifest.xml") -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\FileSystemMetadata.xml") -force -ea 0 

RI @(Join-Path $env:ProgramData "Microsoft\ClickToRun") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramData "Microsoft\office\Heartbeat") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramData "Microsoft\office\FFPackageLocker") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramData "Microsoft\office\ClickToRunPackageLocker") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office 15") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office 16") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\root") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\Office16") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\Office15") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\PackageManifests") -Recurse -force -ea 0 
RI @(Join-Path $env:ProgramFiles "Microsoft Office\PackageSunrisePolicies") -Recurse -force -ea 0 
RI @(Join-Path $env:CommonProgramFiles "microsoft shared\ClickToRun") -Recurse -force -ea 0 

if ($env:ProgramFilesX86) {
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\AppXManifest.xml") -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\FileSystemMetadata.xml") -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\root") -Recurse -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\Office16") -Recurse -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\Office15") -Recurse -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\PackageManifests") -Recurse -force -ea 0 
  RI @(Join-Path $env:ProgramFilesX86 "Microsoft Office\PackageSunrisePolicies") -Recurse -force -ea 0 
}

RI @(Join-Path $env:userprofile "Microsoft Office") -Recurse -force -ea 0 
RI @(Join-Path $env:userprofile "Microsoft Office 15") -Recurse -force -ea 0 
RI @(Join-Path $env:userprofile "Microsoft Office 16") -Recurse -force -ea 0
}
Function RestoreExplorer {
$wmiInfo = gwmi -Query "Select * From Win32_Process Where Name='explorer.exe'"
if (-not $wmiInfo) {
  start "explorer"}
}
Function Uninstall {

"*** -- remove the published component registration for C2R packages"
$Location = (
  "SOFTWARE\Microsoft\Office\ClickToRun",
  "SOFTWARE\Microsoft\Office\16.0\ClickToRun",
  "SOFTWARE\Microsoft\Office\15.0\ClickToRun" )

Foreach ($Loc in $Location) {
  Set-Location "HKLM:\"
  Set-Location $Loc -ea 0
  if (@(Get-Location).Path -ne 'HKLM:\') {
    try {
      $sPkgFld  = $null; $sPkgGuid = $null;
      $sPkgFld  = GPV . -Name PackageFolder
      $sPkgGuid = GPV . -Name PackageGUID
      HandlePakage $sPkgFld $sPkgGuid
    }
    catch {
      $sPkgFld  = $null
      $sPkgGuid = $null
    }
}}

"*** -- delete potential blocking registry keys for msiexec based tasks"
Set-Location "HKLM:\"
RI "HKLM:SOFTWARE\Microsoft\Office\15.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\16.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKLM:SOFTWARE\Microsoft\Office\ClickToRun" -Force -ea 0 -Recurse

Set-Location "HKCU:\"
RI "HKCU:SOFTWARE\Microsoft\Office\15.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\16.0\ClickToRun" -Force -ea 0 -Recurse
RI "HKCU:SOFTWARE\Microsoft\Office\ClickToRun" -Force -ea 0 -Recurse

"*** -- AppV keys"
$hDefKey_List = @(
  "HKCU", "HKLM" )
$sSubKeyName_List = @(
  "SOFTWARE\Microsoft\AppV\ISV",
  "SOFTWARE\Microsoft\AppVISV" )

foreach ($hDefKey in $hDefKey_List) {
  foreach ($sSubKeyName in $sSubKeyName_List) {
    Set-Location "$($hDefKey):\"
    Push-Location "$($hDefKey):$($sSubKeyName)" -ea 0
    if (@(Get-Location).Path -ne "$($hDefKey):\") {
      $arrNames = gi .
      if ($arrNames)  {
        $arrNames.Property | % { 
          $name = GPV . $_
          if ($name -and (
            $Name|IsC2R)) {
              RP . $_ -Force }}}}}}	
	
"*** -- msiexec based uninstall"
try {
  $omsi = Get-MsiProducts }
catch { 
 return }

 if (!($omsi)) { # ! same as -not
   return }
 
$sUninstallCmd = $null
$sMsiProp = "REBOOT=ReallySuppress NOREMOVESPAWN=True"

 $omsi | % {
  $ProductCode   = $_.ProductCode
  $InstallSource = $_.InstallSource

  if (($ProductCode -and ($ProductCode|CheckDelete)) -or (
    $InstallSource -and ($InstallSource|IsC2R))) {
        $sUninstallCmd = "/x $($ProductCode) $($sMsiProp) /q"
	    $proc = start msiexec.exe -Args $sUninstallCmd -Wait -WindowStyle Hidden -ea 0 -PassThru
        "*** -- msIexec $($sUninstallCmd) ,End with value: $($proc.ExitCode)"

 }}
 net stop msiserver *>$null
}
Function RegWipeTypeLib {
$sTLKey = 
"Software\Classes\TypeLib\"

$RegLibs = @(
"\0\Win32\","\0\Win64\","\9\Win32\","\9\Win64\")

$arrTypeLibs = @(
"{000204EF-0000-0000-C000-000000000046}","{000204EF-0000-0000-C000-000000000046}",
"{00020802-0000-0000-C000-000000000046}","{00020813-0000-0000-C000-000000000046}",
"{00020905-0000-0000-C000-000000000046}","{0002123C-0000-0000-C000-000000000046}",
"{00024517-0000-0000-C000-000000000046}","{0002E157-0000-0000-C000-000000000046}",
"{00062FFF-0000-0000-C000-000000000046}","{0006F062-0000-0000-C000-000000000046}",
"{0006F080-0000-0000-C000-000000000046}","{012F24C1-35B0-11D0-BF2D-0000E8D0D146}",
"{06CA6721-CB57-449E-8097-E65B9F543A1A}","{07B06096-5687-4D13-9E32-12B4259C9813}",
"{0A2F2FC4-26E1-457B-83EC-671B8FC4C86D}","{0AF7F3BE-8EA9-4816-889E-3ED22871FE05}",
"{0D452EE1-E08F-101A-852E-02608C4D0BB4}","{0EA692EE-BB50-4E3C-AEF0-356D91732725}",
"{1F8E79BA-9268-4889-ADF3-6D2AABB3C32C}","{2374F0B1-3220-4c71-B702-AF799F31ABB4}",
"{238AA1AC-786F-4C17-BAAB-253670B449B9}","{28DD2950-2D4A-42B5-ABBF-500AA42E7EC1}",
"{2A59CA0A-4F1B-44DF-A216-CB2C831E5870}","{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}",
"{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}","{2F7FC181-292B-11D2-A795-DFAA798E9148}",
"{3120BA9F-4FC8-4A4F-AE1E-02114F421D0A}","{31411197-A502-11D2-BBCA-00C04F8EC294}",
"{3B514091-5A69-4650-87A3-607C4004C8F2}","{47730B06-C23C-4FCA-8E86-42A6A1BC74F4}",
"{49C40DDF-1B04-4868-B3B5-E49F120E4BFA}","{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}",
"{4AFFC9A0-5F99-101B-AF4E-00AA003F0F07}","{4D95030A-A3A9-4C38-ACA8-D323A2267698}",
"{55A108B0-73BB-43db-8C03-1BEF4E3D2FE4}","{56D04F5D-964F-4DBF-8D23-B97989E53418}",
"{5B87B6F0-17C8-11D0-AD41-00A0C90DC8D9}","{66CDD37F-D313-4E81-8C31-4198F3E42C3C}",
"{6911FD67-B842-4E78-80C3-2D48597C2ED0}","{698BB59C-38F1-4CEF-92F9-7E3986E708D3}",
"{6DDCE504-C0DC-4398-8BDB-11545AAA33EF}","{6EFF1177-6974-4ED1-99AB-82905F931B87}",
"{73720002-33A0-11E4-9B9A-00155D152105}","{759EF423-2E8F-4200-ADF0-5B6177224BEE}",
"{76F6F3F5-9937-11D2-93BB-00105A994D2C}","{773F1B9A-35B9-4E95-83A0-A210F2DE3B37}",
"{7D868ACD-1A5D-4A47-A247-F39741353012}","{7E36E7CB-14FB-4F9E-B597-693CE6305ADC}",
"{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}","{8404DD0E-7A27-4399-B1D9-6492B7DD7F7F}",
"{8405D0DF-9FDD-4829-AEAD-8E2B0A18FEA4}","{859D8CF5-7ADE-4DAB-8F7D-AF171643B934}",
"{8E47F3A2-81A4-468E-A401-E1DEBBAE2D8D}","{91493440-5A91-11CF-8700-00AA0060263B}",
"{9A8120F2-2782-47DF-9B62-54F672075EA1}","{9B7C3E2E-25D5-4898-9D85-71CEA8B2B6DD}",
"{9B92EB61-CBC1-11D3-8C2D-00A0CC37B591}","{9D58B963-654A-4625-86AC-345062F53232}",
"{9DCE1FC0-58D3-471B-B069-653CE02DCE88}","{A4D51C5D-F8BF-46CC-92CC-2B34D2D89716}",
"{A717753E-C3A6-4650-9F60-472EB56A7061}","{AA53E405-C36D-478A-BBFF-F359DF962E6D}",
"{AAB9C2AA-6036-4AE1-A41C-A40AB7F39520}","{AB54A09E-1604-4438-9AC7-04BE3E6B0320}",
"{AC0714F2-3D04-11D1-AE7D-00A0C90F26F4}","{AC2DE821-36A2-11CF-8053-00AA006009FA}",
"{B30CDC65-4456-4FAA-93E3-F8A79E21891C}","{B8812619-BDB3-11D0-B19E-00A0C91E29D8}",
"{B9164592-D558-4EE7-8B41-F1C9F66D683A}","{B9AA1F11-F480-4054-A84E-B5D9277E40A8}",
"{BA35B84E-A623-471B-8B09-6D72DD072F25}","{BDEADE33-C265-11D0-BCED-00A0C90AB50F}",
"{BDEADEF0-C265-11D0-BCED-00A0C90AB50F}","{BDEADEF0-C265-11D0-BCED-00A0C90AB50F}",
"{C04E4E5E-89E6-43C0-92BD-D3F2C7FBA5C4}","{C3D19104-7A67-4EB0-B459-D5B2E734D430}",
"{C78F486B-F679-4af5-9166-4E4D7EA1CEFC}","{CA973FCA-E9C3-4B24-B864-7218FC1DA7BA}",
"{CBA4EBC4-0C04-468d-9F69-EF3FEED03236}","{CBBC4772-C9A4-4FE8-B34B-5EFBD68F8E27}",
"{CD2194AA-11BE-4EFD-97A6-74C39C6508FF}","{E0B12BAE-FC67-446C-AAE8-4FA1F00153A7}",
"{E985809A-84A6-4F35-86D6-9B52119AB9D7}","{ECD5307E-4419-43CF-8BDA-C9946AC375CF}",
"{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}","{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}",
"{EDCD5812-6A06-43C3-AFAC-46EF5D14E22C}","{EDDCFF16-3AEE-4883-BD91-0F3978640DFB}",
"{EE9CFA8C-F997-4221-BE2F-85A5F603218F}","{F2A7EE29-8BF6-4a6d-83F1-098E366C709C}",
"{F3685D71-1FC6-4CBD-B244-E60D8C89990B}")

    foreach ($tl in $arrTypeLibs) {
  
      Set-Location "HKLM:\"
      $sKey = "HKLM:" + $sTLKey + $tl

      Set-Location "HKLM:\"
      Push-Location $sKey -ea 0
      if (@(Get-Location).Path -eq 'HKLM:\') {
        continue
      }

      $children   = GCI .
      $fCanDelete = $false

      if (-not $children) {
        Set-Location "HKLM:\"
        Push-Location "HKLM:$($sTLKey)" -ea 0
        if (@(Get-Location).Path -ne 'HKLM:\') {
          RI $tl -Recurse -Force }
        continue
      }
  
      foreach ($K in $children) {
    
        $sTLVerKey = $sKey + "\" + $K.PSChildName
        $PSChildName = GCI $K.PSChildName -ea 0
        if ($PSChildName) {
          $fCanDelete = $true }
    
        Set-Location "HKLM:\"
        Push-Location $sKey -ea 0
        if (@(Get-Location).Path -eq 'HKLM:\') {
          continue }
    
        $RegLibs | % {
          Set-Location "HKLM:\"
          Push-Location "$($sTLVerKey)$($_)" -ea 0
          if (@(Get-Location).Path -ne 'HKLM:\') {
            try {
              $Default = gpv . -Name '(Default)' -ea 0 }
            catch {}
            if ($Default -and (
              [IO.FILE]::Exists($Default))) {
                $fCanDelete = $false }}}

        if ($fCanDelete) {
          Set-Location "HKLM:\"
          Push-Location $sKey -ea 0
          if (@(Get-Location).Path -ne 'HKLM:\') {
	      RI $K.PSChildName -Recurse -Force }}
      }
    }
}
Function CleanOSPP {

    $oProductInstances = gwmi -Query "SELECT * FROM SoftwareLicensingProduct WHERE (ApplicationId = '0ff1ce15-a989-479d-af46-f275c6370663' AND PartialProductKey Is not NULL)" -ea 0
    if ($oProductInstances) {
     $oProductInstances | % {
       $_.UninstallProductKey()|Out-Null}
    }

    Set-Location "HKLM:\"
    $OSPP = (Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\OfficeSoftwareProtectionPlatform" -ea 0).Path
    if ($OSPP) {
        <# Write-Output "Found Office Software Protection installed, cleaning" #>
        UninstallLicenses($OSPP + "osppc.dll")
    }
    UninstallLicenses("sppc.dll")
}
Function ClearVNextLicCache {

$Licenses = Join-Path $ENV:localappdata "Microsoft\Office\Licenses"
    if (Test-Path $Licenses) {
      Set-Location "$($env:SystemDrive)\"
      RI $Licenses -Recurse -Force -ea 0 }
}
Function HandlePakage {
  param (
   [parameter(Mandatory=$True)]
   [string]$sPkgFldr,

   [parameter(Mandatory=$True)]
   [string]$sPkgGuid
  )

  $RootPath =
    Join-Path $sPkgFldr "\root"
  $IntegrationPath =
    Join-Path $sPkgFldr "\root\Integration"
  $Integrator =
    Join-Path $sPkgFldr "\root\Integration\Integrator.exe"
  $Integrator_ =
    "$env:ProgramData\Microsoft\ClickToRun\{$sPkgGuid}\integrator.exe"

  if (-not (
      Test-Path ($IntegrationPath ))) {
        return }
  
  Set-Location 'c:\'
  Push-Location $RootPath

  #Remove `Root`->`Integration\C2RManifest*.xml`
  if (@(Get-Location).Path -ne 'c:\') {
    RI .\Integration\ -Filter "C2RManifest*.xml" -Recurse -Force -ea 0
  }
  
  if ([IO.FILE]::Exists(
    $Integrator)) {
      $Args = "/U /Extension PackageRoot=""$($RootPath)"" PackageGUID=""$($sPkgGuid)"""
      $Proc = start $Integrator -arg $Args -Wait -WindowStyle Hidden -PassThru -ea 0
	  "*** -- Uninstall ID: $sPkgGuid with Full Args, returned with value:$($Proc.ExitCode)"
      $Args = "/U"
      $Proc = start $Integrator -arg $Args -Wait -WindowStyle Hidden -PassThru  -ea 0
	  "*** -- Uninstall ID: $sPkgGuid with Minimum Args, returned with value:$($Proc.ExitCode)" }

  if ([IO.FILE]::Exists(
    $Integrator_)) {
      $Args = "/U /Extension PackageRoot=""$($RootPath)"" PackageGUID=""$($sPkgGuid)"""
      $Proc = start $Integrator_ -arg $Args -Wait -WindowStyle Hidden -PassThru -ea 0
	  "*** -- Uninstall ID: $sPkgGuid with Full Args, returned with value:$($Proc.ExitCode)"
      $Args = "/U"
      $Proc = start $Integrator_ -arg $Args -Wait -WindowStyle Hidden -PassThru  -ea 0
	  "*** -- Uninstall ID: $sPkgGuid with Minimum Args, returned with value:$($Proc.ExitCode)" }
}

# ---------------------- #
# Begin of main function #
# ---------------------- #

"*** $(Get-Date -Format hh:mm:ss): Load HKCR Hive"
if ($null -eq (Get-PSDrive HKCR -ea 0)) {
    New-PSDrive HKCR Registry HKEY_CLASSES_ROOT -ea 1 | Out-Null }

"*** $(Get-Date -Format hh:mm:ss): Clean OSPP"
CleanOSPP

"*** $(Get-Date -Format hh:mm:ss): Clean vNext Licenses"
ClearVNextLicCache

"*** $(Get-Date -Format hh:mm:ss): End running processes"
ClearShellIntegrationReg
CloseOfficeApps

"*** $(Get-Date -Format hh:mm:ss): Clean Scheduler tasks"
DelSchtasks

"*** $(Get-Date -Format hh:mm:ss): Clean Office shortcuts"
CleanShortcuts -sFolder "$env:AllusersProfile"
CleanShortcuts -sFolder "$env:SystemDrive\Users"

"*** $(Get-Date -Format hh:mm:ss): Remove Office C2R / O365"
Uninstall

"*** $(Get-Date -Format hh:mm:ss): call odt based uninstall"
UninstallOfficeC2R

"*** $(Get-Date -Format hh:mm:ss): CleanUp"
FileWipe
RegWipe

"*** $(Get-Date -Format hh:mm:ss): Ensure Explorer runs"
RestoreExplorer

"*** $(Get-Date -Format hh:mm:ss): Un-Load HKCR Hive"
Set-Location "HKLM:\"
Remove-PSDrive -Name HKCR -ea 0 | Out-Null

write-host "Begin: $($Start_Time), End: $(Get-Date -Format hh:mm:ss)"
Write-Host
timeout 3 *>$null

# SIG # Begin signature block
# MIIFoQYJKoZIhvcNAQcCoIIFkjCCBY4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUf4IKws5DOH4vWhB+RKOZwWmU
# 6pKgggM2MIIDMjCCAhqgAwIBAgIQSe49BypwLKhOHkzMeRKLBzANBgkqhkiG9w0B
# AQsFADAgMR4wHAYDVQQDDBVhZG1pbkBvZmZpY2VydG9vbC5vcmcwHhcNMjQwMTA2
# MTYxMjI3WhcNMzAwMTA2MTYyMjI3WjAgMR4wHAYDVQQDDBVhZG1pbkBvZmZpY2Vy
# dG9vbC5vcmcwggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQDLZq+Rmrz4
# wwNvgAZVzvbOmj1RlUll7htG/vIJurDabWNvIbYBxycLrzEAJKeuuO8TTtodlhCF
# kvCtzO2gU47wKwqoIK9p5orB9f0xasuxtu7EeIRvXZLpBjKQ20Fnzed6peoPupEb
# 5+2FIjAbM3ErtSbmC7XDhSLhAheV8+Urio/vv7zhiI0JYsfKtcZnbFBG8h5WOoYS
# k7vEF6nW4OleuM6oGuprq7OWDYGLa9sarX8mjNu0CPDgvxoE6vAiOY6lXgT9GoSn
# EOgpn8OOhpBp9ERPzP6Qq6qetl/+wYGkYbQGz7v6fPDQ4ATnGFIfc9G+qICE8iZs
# TV+bgDYjyMUJAgMBAAGjaDBmMA4GA1UdDwEB/wQEAwIHgDATBgNVHSUEDDAKBggr
# BgEFBQcDAzAgBgNVHREEGTAXghVhZG1pbkBvZmZpY2VydG9vbC5vcmcwHQYDVR0O
# BBYEFDIRoZpOPb0eh3mSqlUpHSSgioiqMA0GCSqGSIb3DQEBCwUAA4IBAQCe7S09
# 5VqCa9rw0s6FdNqvHOftRgRf1or04i7ZuV3jffN/rValXj8O7GtTRQ9ZFhGInjZ/
# 5UVyLPnhVqVwWzNpqFgTL5/Y0joFQ99GQfv1f5UUt6U4jNjjSTZNdVa3C9iwV4IQ
# jaRhGEQqsvqsOadezbX9jlIpXBKxmua70/cUj8Ub0UBT+jrt3ztqsX/Ei/wrorbh
# 8qS1rgYmi493hgQgKxSG/7tZ5PvbljEO5KPEMagKF6u4XX1B7Mz0DQAJcFUnTsNy
# D/Tj8nc03aYnF8NRkUyRYPhbIgpiY9e7/ivBY+4gF20ONc1Cy8+zqgSn17mF1QTD
# TOzL7jtV+7ROPKxOMYIB1TCCAdECAQEwNDAgMR4wHAYDVQQDDBVhZG1pbkBvZmZp
# Y2VydG9vbC5vcmcCEEnuPQcqcCyoTh5MzHkSiwcwCQYFKw4DAhoFAKB4MBgGCisG
# AQQBgjcCAQwxCjAIoAKAAKECgAAwGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQw
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFHcI
# 22qZQgFmaIlE/kx4cWZpd2P+MA0GCSqGSIb3DQEBAQUABIIBABHWwTfC6E3Erzd/
# Fj48ZJ1it3qX5s34s5xbUyPBoKElXPMPjP79GrgYGaFnul6Cq/DZEJg9JQcAlgTG
# +UMj436H1L5zcMhFxtqhop+nop8XVRcMxysLCmrDaCpP8R7nfccMVzkFvZJJz5Tq
# NC9rPwbzc8ygsMxg3mGONLPiJzSBzYpGbST7/j5KyjcbJB88wTItdqpBnycm+MwF
# RexMMpwmsLDyuXfi/73tviMTGYu2ZDL45kcMpQbD4MToMeJzaBFnv9AE2Y+DkFmR
# ixedTe39ZDWHbOQ1w4G+hqP7lKflCjg21pdAYzr+z9qYtFnUYflcTiVb5I8IYxMS
# 0TQ5Axs=
# SIG # End signature block
