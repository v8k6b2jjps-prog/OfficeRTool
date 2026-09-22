#requires -Version 5.1
#requires -RunAsAdministrator

# Manually Activate Windows
# https://massgrave.dev/manual_hwid_activation
# https://massgrave.dev/manual_kms38_activation
# https://massgrave.dev/manual_ohook_activation

# Define the KeyBlock function to create a hash table
function Get-KeyBlock {
    
$inputText = @'
O365BusinessRetail,Y9NF9-M2QWD-FF6RJ-QJW36-RRF2T
O365EduCloudRetail,W62NQ-267QR-RTF74-PF2MH-JQMTH
O365HomePremRetail,3NMDC-G7C3W-68RGP-CB4MH-4CXCH
O365ProPlusRetail,H8DN8-Y2YP3-CR9JT-DHDR9-C7GP3
O365SmallBusPremRetail,2QCNB-RMDKJ-GC8PB-7QGQV-7QTQJ
O365AppsBasicRetail,3HYJN-9KG99-F8VG9-V3DT8-JFMHV
AccessRetail,WHK4N-YQGHB-XWXCC-G3HYC-6JF94
AccessRuntimeRetail,RNB7V-P48F4-3FYY6-2P3R3-63BQV
AccessVolume,JJ2Y4-N8KM3-Y8KY3-Y22FR-R3KVK
ExcelRetail,RKJBN-VWTM2-BDKXX-RKQFD-JTYQ2
ExcelVolume,FVGNR-X82B2-6PRJM-YT4W7-8HV36
HomeBusinessPipcRetail,2WQNF-GBK4B-XVG6F-BBMX7-M4F2Y
HomeBusinessRetail,HM6FM-NVF78-KV9PM-F36B8-D9MXD
HomeStudentARMRetail,PBQPJ-NC22K-69MXD-KWMRF-WFG77
HomeStudentPlusARMRetail,6F2NY-7RTX4-MD9KM-TJ43H-94TBT
HomeStudentRetail,PNPRV-F2627-Q8JVC-3DGR9-WTYRK
HomeStudentVNextRetail,YWD4R-CNKVT-VG8VJ-9333B-RC3B8
MondoRetail,VNWHF-FKFBW-Q2RGD-HYHWF-R3HH2
MondoVolume,FMTQQ-84NR8-2744R-MXF4P-PGYR3
OneNoteFreeRetail,XYNTG-R96FY-369HX-YFPHY-F9CPM
OneNoteRetail,FXF6F-CNC26-W643C-K6KB7-6XXW3
OneNoteVolume,9TYVN-D76HK-BVMWT-Y7G88-9TPPV
OutlookRetail,7N4KG-P2QDH-86V9C-DJFVF-369W9
OutlookVolume,7QPNR-3HFDG-YP6T9-JQCKQ-KKXXC
PersonalPipcRetail,9CYB3-NFMRW-YFDG6-XC7TF-BY36J
PersonalRetail,FT7VF-XBN92-HPDJV-RHMBY-6VKBF
PowerPointRetail,N7GCB-WQT7K-QRHWG-TTPYD-7T9XF
PowerPointVolume,X3RT9-NDG64-VMK2M-KQ6XY-DPFGV
ProPlusRetail,GM43N-F742Q-6JDDK-M622J-J8GDV
ProPlusVolume,FNVK8-8DVCJ-F7X3J-KGVQB-RC2QY
ProfessionalPipcRetail,CF9DD-6CNW2-BJWJQ-CVCFX-Y7TXD
ProfessionalRetail,NXFTK-YD9Y7-X9MMJ-9BWM6-J2QVH
ProjectProRetail,WPY8N-PDPY4-FC7TF-KMP7P-KWYFY
ProjectProVolume,PKC3N-8F99H-28MVY-J4RYY-CWGDH
ProjectProXVolume,JBNPH-YF2F7-Q9Y29-86CTG-C9YGV
ProjectStdRetail,NTHQT-VKK6W-BRB87-HV346-Y96W8
ProjectStdVolume,4TGWV-6N9P6-G2H8Y-2HWKB-B4G93
ProjectStdXVolume,N3W2Q-69MBT-27RD9-BH8V3-JT2C8
PublisherRetail,WKWND-X6G9G-CDMTV-CPGYJ-6MVBF
PublisherVolume,9QVN2-PXXRX-8V4W8-Q7926-TJGD8
SkypeServiceBypassRetail,6MDN4-WF3FV-4WH3Q-W699V-RGCMY
SkypeforBusinessEntryRetail,4N4D8-3J7Y3-YYW7C-73HD2-V8RHY
SkypeforBusinessRetail,PBJ79-77NY4-VRGFG-Y8WYC-CKCRC
SkypeforBusinessVolume,DMTCJ-KNRKR-JV8TQ-V2CR2-VFTFH
StandardRetail,2FPWN-4H6CM-KD8QQ-8HCHC-P9XYW
StandardVolume,WHGMQ-JNMGT-MDQVF-WDR69-KQBWC
VisioProRetail,NVK2G-2MY4G-7JX2P-7D6F2-VFQBR
VisioProVolume,NRKT9-C8GP2-XDYXQ-YW72K-MG92B
VisioProXVolume,G98Q2-B6N77-CFH9J-K824G-XQCC4
VisioStdRetail,NCRB7-VP48F-43FYY-62P3R-367WK
VisioStdVolume,XNCJB-YY883-JRW64-DPXMX-JXCR6
VisioStdXVolume,B2HTN-JPH8C-J6Y6V-HCHKB-43MGT
WordRetail,P8K82-NQ7GG-JKY8T-6VHVY-88GGD
WordVolume,YHMWC-YN6V9-WJPXD-3WQKP-TMVCV
Access2019Retail,WRYJ6-G3NP7-7VH94-8X7KP-JB7HC
Access2019Volume,6FWHX-NKYXK-BW34Q-7XC9F-Q9PX7
AccessRuntime2019Retail,FGQNJ-JWJCG-7Q8MG-RMRGJ-9TQVF
Excel2019Retail,KBPNW-64CMM-8KWCB-23F44-8B7HM
Excel2019Volume,8NT4X-GQMCK-62X4P-TW6QP-YKPYF
HomeBusiness2019Retail,QBN2Y-9B284-9KW78-K48PB-R62YT
HomeStudentARM2019Retail,DJTNY-4HDWM-TDWB2-8PWC2-W2RRT
HomeStudentPlusARM2019Retail,NM8WT-CFHB2-QBGXK-J8W6J-GVK8F
HomeStudent2019Retail,XNWPM-32XQC-Y7QJC-QGGBV-YY7JK
Outlook2019Retail,WR43D-NMWQQ-HCQR2-VKXDR-37B7H
Outlook2019Volume,RN3QB-GT6D7-YB3VH-F3RPB-3GQYB
Personal2019Retail,NMBY8-V3CV7-BX6K6-2922Y-43M7T
PowerPoint2019Retail,HN27K-JHJ8R-7T7KK-WJYC3-FM7MM
PowerPoint2019Volume,29GNM-VM33V-WR23K-HG2DT-KTQYR
ProPlus2019Retail,BN4XJ-R9DYY-96W48-YK8DM-MY7PY
ProPlus2019Volume,T8YBN-4YV3X-KK24Q-QXBD7-T3C63
Professional2019Retail,9NXDK-MRY98-2VJV8-GF73J-TQ9FK
ProjectPro2019Retail,JDTNC-PP77T-T9H2W-G4J2J-VH8JK
ProjectPro2019Volume,TBXBD-FNWKJ-WRHBD-KBPHH-XD9F2
ProjectStd2019Retail,R3JNT-8PBDP-MTWCK-VD2V8-HMKF9
ProjectStd2019Volume,RBRFX-MQNDJ-4XFHF-7QVDR-JHXGC
Publisher2019Retail,4QC36-NW3YH-D2Y9D-RJPC7-VVB9D
Publisher2019Volume,K8F2D-NBM32-BF26V-YCKFJ-29Y9W
SkypeforBusiness2019Retail,JBDKF-6NCD6-49K3G-2TV79-BKP73
SkypeforBusiness2019Volume,9MNQ7-YPQ3B-6WJXM-G83T3-CBBDK
SkypeforBusinessEntry2019Retail,N9722-BV9H6-WTJTT-FPB93-978MK
Standard2019Retail,NDGVM-MD27H-2XHVC-KDDX2-YKP74
Standard2019Volume,NT3V6-XMBK7-Q66MF-VMKR4-FC33M
VisioPro2019Retail,2NWVW-QGF4T-9CPMB-WYDQ9-7XP79
VisioPro2019Volume,33YF4-GNCQ3-J6GDM-J67P3-FM7QP
VisioStd2019Retail,263WK-3N797-7R437-28BKG-3V8M8
VisioStd2019Volume,BGNHX-QTPRJ-F9C9G-R8QQG-8T27F
Word2019Retail,JXR8H-NJ3MK-X66W8-78CWD-QRVR2
Word2019Volume,9F36R-PNVHH-3DXGQ-7CD2H-R9D3V
Access2021Retail,P286B-N3XYP-36QRQ-29CMP-RVX9M
AccessRuntime2021Retail,MNX9D-PB834-VCGY2-K2RW2-2DP3D
Access2021Volume,JBH3N-P97FP-FRTJD-MGK2C-VFWG6
Excel2021Retail,V6QFB-7N7G9-PF7W9-M8FQM-MY8G9
Excel2021Volume,WNYR4-KMR9H-KVC8W-7HJ8B-K79DQ
HomeBusiness2021Retail,JM99N-4MMD8-DQCGJ-VMYFY-R63YK
HomeStudent2021Retail,N3CWD-38XVH-KRX2Y-YRP74-6RBB2
OneNoteFree2021Retail,CNM3W-V94GB-QJQHH-BDQ3J-33Y8H
OneNote2021Retail,NB2TQ-3Y79C-77C6M-QMY7H-7QY8P
OneNote2021Volume,THNKC-KFR6C-Y86Q9-W8CB3-GF7PD
Outlook2021Retail,4NCWR-9V92Y-34VB2-RPTHR-YTGR7
Outlook2021Volume,JQ9MJ-QYN6B-67PX9-GYFVY-QJ6TB
Personal2021Retail,RRRYB-DN749-GCPW4-9H6VK-HCHPT
PowerPoint2021Retail,3KXXQ-PVN2C-8P7YY-HCV88-GVM96
PowerPoint2021Volume,39G2N-3BD9C-C4XCM-BD4QG-FVYDY
ProPlus2021Retail,8WXTP-MN628-KY44G-VJWCK-C7PCF
ProPlus2021Volume,RNHJY-DTFXW-HW9F8-4982D-MD2CW
ProPlusSPLA2021Volume,JRJNJ-33M7C-R73X3-P9XF7-R9F6M
Professional2021Retail,DJPHV-NCJV6-GWPT6-K26JX-C7PBG
ProjectPro2021Retail,QKHNX-M9GGH-T3QMW-YPK4Q-QRWMV
ProjectPro2021Volume,HVC34-CVNPG-RVCMT-X2JRF-CR7RK
ProjectStd2021Retail,2B96V-X9NJY-WFBRC-Q8MP2-7CHRR
ProjectStd2021Volume,3CNQX-T34TY-99RH4-C4YD2-KW6WH
Publisher2021Retail,CDNFG-77T8D-VKQJX-B7KT3-KK28V
Publisher2021Volume,2KXJH-3NHTW-RDBPX-QFRXJ-MTGXF
SkypeforBusiness2021Retail,DVBXN-HFT43-CVPRQ-J89TF-VMMHG
SkypeforBusiness2021Volume,R3FCY-NHGC7-CBPVP-8Q934-YTGXG
Standard2021Retail,HXNXB-J4JGM-TCF44-2X2CV-FJVVH
Standard2021Volume,2CJN4-C9XK2-HFPQ6-YH498-82TXH
StandardSPLA2021Volume,BQWDW-NJ9YF-P7Y79-H6DCT-MKQ9C
VisioPro2021Retail,T6P26-NJVBR-76BK8-WBCDY-TX3BC
VisioPro2021Volume,JNKBX-MH9P4-K8YYV-8CG2Y-VQ2C8
VisioStd2021Retail,89NYY-KB93R-7X22F-93QDF-DJ6YM
VisioStd2021Volume,BW43B-4PNFP-V637F-23TR2-J47TX
Word2021Retail,VNCC4-CJQVK-BKX34-77Y8H-CYXMR
Word2021Volume,BJG97-NW3GM-8QQQ7-FH76G-686XM
Access2024Retail,P6NMW-JMTRC-R6MQ6-HH3F2-BTHKB
Access2024Volume,CXNJT-98HPP-92HX7-MX6GY-2PVFR
Excel2024Retail,82CNJ-W82TW-BY23W-BVJ6W-W48GP
Excel2024Volume,7Y287-9N2KC-8MRR3-BKY82-2DQRV
Home2024Retail,N69X7-73KPT-899FD-P8HQ4-QGTP4
HomeBusiness2024Retail,PRKQM-YNPQR-77QT6-328D7-BD223
Outlook2024Retail,2CFK4-N44KG-7XG89-CWDG6-P7P27
Outlook2024Volume,NQPXP-WVB87-H3MMB-FYBW2-9QFPB
PowerPoint2024Retail,CT2KT-GTNWH-9HFGW-J2PWJ-XW7KJ
PowerPoint2024Volume,RRXFN-JJ26R-RVWD2-V7WMP-27PWQ
ProjectPro2024Retail,GNJ6P-Y4RBM-C32WW-2VJKJ-MTHKK
ProjectPro2024Volume,WNFMR-HK4R7-7FJVM-VQ3JC-76HF6
ProjectStd2024Retail,C2PNM-2GQFC-CY3XR-WXCP4-GX3XM
ProjectStd2024Volume,F2VNW-MW8TT-K622Q-4D96H-PWJ8X
ProPlus2024Retail,VWCNX-7FKBD-FHJYG-XBR4B-88KC6
ProPlus2024Volume,4YV2J-VNG7W-YGTP3-443TK-TF8CP
SkypeforBusiness2024Volume,XKRBW-KN2FF-G8CKY-HXVG6-FVY2V
Standard2024Volume,GVG6N-6WCHH-K2MVP-RQ78V-3J7GJ
VisioPro2024Retail,HGRBX-N68QF-6DY8J-CGX4W-XW7KP
VisioPro2024Volume,GBNHB-B2G3Q-G42YB-3MFC2-7CJCX
VisioStd2024Retail,VBXPJ-38NR3-C4DKF-C8RT7-RGHKQ
VisioStd2024Volume,YNFTY-63K7P-FKHXK-28YYT-D32XB
Word2024Retail,XN33R-RP676-GMY2F-T3MH7-GCVKR
Word2024Volume,WD8CQ-6KNQM-8W2CX-2RT63-KK3TP
'@

# Initialize an empty hash table
$hashTable = @{}
    
# Process each line of input
$inputText -split "`n" | ForEach-Object {
    $parts = $_ -split ","
        
    if ($parts.Length -eq 2) {
        # Add the Product as the key and the GenericKey as the value
        $hashTable[$parts[0].Trim()] = $parts[1].Trim()
    }
}
    
return $hashTable
}

# Install function
function Install {
# Ensure WMI service is available before running the script
try {
    $wmiService = Get-WmiObject -Class SoftwareLicensingService -ErrorAction Stop
} catch {
    Write-Host "Error accessing WMI service: $_"
    Start-Sleep 4
    return
}

# Define registry path and value
$regPath = "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration"

# Get ProductReleaseIds registry value and process it
try {
    $productReleaseIds = (Get-ItemProperty -Path $regPath).ProductReleaseIds
} catch {
    Write-Host "Error accessing registry: $_"
    Start-Sleep 4
    return
}    
if ($productReleaseIds) {
    $keyBlock = Get-KeyBlock
    
    $productReleaseIds -split ',' | ForEach-Object {
        $productKey = $_.Trim()
        if ($keyBlock.ContainsKey($productKey)) {
            try {
                $wmiService.InstallProductKey($keyBlock[$productKey]) | Out-Null
                Write-Host "Installed key for: ${productKey}"
            } catch {
                Write-Host "Failed to install key for ${productKey}: $_"
                Start-Sleep 4
                return
            }
        }
    }
} else {
    Write-Host "ProductReleaseIds not found."
}

$base64 = @'
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
'@

# Decode the Base64 string back into a byte array
$compressedDllBytes = [Convert]::FromBase64String($base64)

# Step 2: Create a MemoryStream from the byte array
$compressedStream = [System.IO.MemoryStream]::new($compressedDllBytes)

# Step 3: Decompress the byte array using GZip
$gzipStream = New-Object System.IO.Compression.GZipStream($compressedStream, [System.IO.Compression.CompressionMode]::Decompress)
$decompressedStream = New-Object System.IO.MemoryStream

# Copy the decompressed data to a new MemoryStream
$gzipStream.CopyTo($decompressedStream)
$gzipStream.Close()

# Step 4: Get the decompressed byte array
$decompressedDllBytes = $decompressedStream.ToArray()
    
# Define the output file path with ProgramFiles environment variable
$sppc = [System.IO.Path]::Combine($env:ProgramFiles, "Microsoft Office\root\vfs\System\sppc.dll")

# Define paths for symbolic link and target
$sppcs = [System.IO.Path]::Combine($env:ProgramFiles, "Microsoft Office\root\vfs\System\sppcs.dll")
$System32 = [System.IO.Path]::Combine($env:windir, "System32\sppc.dll")

# Step 1: Check if the symbolic link exists and remove it if necessary
if (Test-Path -Path $sppcs) {
    Write-Host "Symbolic link already exists at $sppcs. Attempting to remove..."

    try {
        # Remove the existing symbolic link
        Remove-Item -Path $sppcs -Force
        Write-Host "Existing symbolic link removed successfully."
    } catch {
        Write-Host "Failed to remove existing symbolic link: $_"
    }
} else {
    Write-Host "No symbolic link found at $sppcs."
}

try {
    # Attempt to write byte array to the file
    [System.IO.File]::WriteAllBytes($sppc, $decompressedDllBytes)
    Write-Host "Byte array written successfully to $sppc."
} 
catch {
    Write-Host "Failed to write byte array to ${sppc}: $_"

    # Inner try-catch to handle the case where the file is in use
    try {
        Write-Host "File is in use or locked. Attempting to move it to a temp file..."

        # Generate a random name for the temporary file in the temp folder
        $tempDir = [System.IO.Path]::GetTempPath()
        $tempFileName = [System.IO.Path]::Combine($tempDir, [System.Guid]::NewGuid().ToString() + ".bak")

        # Move the file to the temp location with a random name
        Move-Item -Path $sppc -Destination $tempFileName -Force
        Write-Host "Moved file to temporary location: $tempFileName"

        # Retry the write operation after moving the file
        [System.IO.File]::WriteAllBytes($sppc, $decompressedDllBytes)
        Write-Host "Byte array written successfully to $sppc after moving the file."

    } catch {
        Write-Host "Failed to move the file or retry the write operation: $_"
    }
}

# Step 3: Check if the symbolic link exists and create it if necessary
try {
    if (-not (Test-Path -Path $sppcs)) {
        # Create symbolic link only if it doesn't already exist
        New-Item -Path $sppcs -ItemType SymbolicLink -Target $System32 | Out-Null
        Write-Host "Symbolic link created successfully at $sppcs."
    } else {
        Write-Host "Symbolic link already exists at $sppcs."
    }
} catch {
    Write-Host "Failed to create symbolic link at ${sppcs}: $_"
}
	
# Define the target registry key path
$RegPath = "HKCU:\Software\Microsoft\Office\16.0\Common\Licensing\Resiliency"

# Define the value name and data
$ValueName = "TimeOfLastHeartbeatFailure"
$ValueData = "2040-01-01T00:00:00Z"

# Check if the registry key exists. If not, create it.
# The -Force parameter on New-Item ensures the full path is created if necessary.
if (-not (Test-Path -Path $RegPath)) {
	Write-host "Registry key '$RegPath' not found. Creating it."
	# Use -Force to create the key and any missing parent keys
	# Out-Null is used to suppress the output object from New-Item
	New-Item -Path $RegPath -Force | Out-Null
}

# Set the registry value within the existing (or newly created) key.
# The -Force parameter on Set-ItemProperty ensures the value is created if it doesn't exist
# or updated if it does exist.
Write-host "Setting registry value '$ValueName' at '$RegPath'."
Set-ItemProperty -Path $RegPath -Name $ValueName -Value $ValueData -Type String -Force
}

# Remove function
function Remove {
    
# Remove the symbolic link if it exists
$sppcs = [System.IO.Path]::Combine($env:ProgramFiles, "Microsoft Office\root\vfs\System\sppcs.dll")
if (Test-Path -Path $sppcs) {
    try {
        Remove-Item -Path $sppcs -Force
        Write-Host "Symbolic link '$sppcs' removed successfully."
    } catch {
        Write-Host "Failed to remove symbolic link '$sppcs': $_"
    }
} else {
    Write-Host "No symbolic link found at '$sppcs'."
}

# Remove the actual DLL file if it exists
$sppc = [System.IO.Path]::Combine($env:ProgramFiles, "Microsoft Office\root\vfs\System\sppc.dll")
if (Test-Path -Path $sppc) {
    try {
        # Try to remove the file and handle any errors if they occur
        Remove-Item -Path $sppc -Force -ErrorAction Stop
        Write-Host "DLL file '$sppc' removed successfully."
    } catch {
        Write-Host "Failed to remove DLL file '$sppc': $_"
            
        # If removal failed, try to move the file to a temporary location
        try {
            # Generate a random name for the file in the temp directory
            $tempDir = [System.IO.Path]::GetTempPath()
            $tempFileName = [System.IO.Path]::Combine($tempDir, [System.Guid]::NewGuid().ToString() + ".bak")
            
            # Attempt to move the file to the temp directory with a random name
            Move-Item -Path $sppc -Destination $tempFileName -Force -ErrorAction Stop
            Write-Host "DLL file moved to Temp folder."
        } catch {
            Write-Host "Failed to move DLL file '$sppc' to temporary location: $_"
        }
    }
} else {
    Write-Host "No DLL file found at '$sppc'."
}
}

# Clear the screen
cls
Write-Host
Write-Host "Welcome to the oHook DLL Installtion Script" -ForegroundColor Cyan
Write-Host "-------------------------------------------"
Write-Host

# Prompt the user for action (I for Install, R for Remove)
$action = Read-Host "Do you want to Install (I) or Remove (R)? (Enter 'I' or 'R')"

# Normalize the input to uppercase for better consistency
$action = $action.ToUpper()

Write-Host

# Run the appropriate function based on user input
switch ($action) {
    'I' {
        Install
        break
    }
    'R' {
        Remove
        break
    }
    default {
        Write-Host "Invalid choice. Please enter either 'I' for Install or 'R' for Remove." -ForegroundColor Red
        break
    }
}

Write-Host "--------------------------------------"
Write-Host "Script execution completed."
# SIG # Begin signature block
# MIIFoQYJKoZIhvcNAQcCoIIFkjCCBY4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUd+2t9kdz8hExM8OzI0OC2Q7e
# s9agggM2MIIDMjCCAhqgAwIBAgIQSe49BypwLKhOHkzMeRKLBzANBgkqhkiG9w0B
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
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFGxx
# 8o97mzdFWj5G3KXj2aubeVIkMA0GCSqGSIb3DQEBAQUABIIBADB4/2FAv7tDb62n
# CERtSFJbe9axLCmndf8bOzrbMTZg0huCAbreuRM27SOzHj3HsQs5sJo5a6/A+KzA
# brG24mDtJgjHoO+q2sktSQpBr+mACWU3kABmKOysJUsG6i8zweZElQUgD2NYyXl8
# 9SIc/hgo5FpQ00q9t7o+HAfxnthat7UMtpA8LTkX7p/HIx8MZRPsBxXPeAB8ToSL
# PaqNr5jgDvGWHIlmWSc2Y9E7xkeZFwVnx4QWMgzh+0G0AaZc/WF5dxGO2GFeHkNO
# TVEVlENj781gNBiq+y2y1VaT8zkl8zoVPqw20ZaS75abjHUcGjc7B6ex+Atf+DaH
# V9D9g1M=
# SIG # End signature block
