
<#
version      : 4.0
Release Date : 27-02-2024
Made By      : Dark Dinosaur, MDL
#>

Function Convert-To-System {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $NAME
   )
  
  switch ($NAME){
   "ARM"      {return "ARM"}
   "CHPE"     {return "CHPE"}
   "Win7"     {return "7.0"}
   "Win8"     {return "8.0"}
   "Win8.0"   {return "8.0"}
   "Win8.1"   {return "8.1"}
   "Default"  {return "10.0"}
   "RDX Test" {return "RDX"}
  }
  return "Null"
}
Function Convert-To-Channel {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
     [string] $FFN
   )
  
  switch ($FFN){
   "492350f6-3a01-4f97-b9c0-c7c6ddf67d60" {return "Current"}
   "64256afe-f5d9-4f86-8936-8840a6a4f5be" {return "CurrentPreview"}
   "5440fd1f-7ecb-4221-8110-145efaa6372f" {return "BetaChannel"}
   "55336b82-a18d-4dd6-b5f6-9e5095c314a6" {return "MonthlyEnterprise"}
   "7ffbc6bf-bc32-4f92-8982-f9dd17fd3114" {return "SemiAnnual"}
   "b8f9b850-328d-4355-9145-c59439a0c4cf" {return "SemiAnnualPreview"}
   "f2e724c1-748f-4b47-8fb8-8e0d210e9208" {return "PerpetualVL2019"}
   "5030841d-c919-4594-8d2d-84ae4f96e58e" {return "PerpetualVL2021"}
   "7983BAC0-E531-40CF-BE00-FD24FE66619C" {return "PerpetualVL2024"}
   "ea4a4090-de26-49d7-93c1-91bff9e53fc3" {return "DogfoodDevMain"}
   "f3260cf1-a92c-4c75-b02e-d64c0a86a968" {return "DogfoodCC"}
   "c4a7726f-06ea-48e2-a13a-9d78849eb706" {return "DogfoodDCEXT"}
   "834504cc-dc55-4c6d-9e71-e024d0253f6d" {return "DogfoodFRDC"}
   "5462eee5-1e97-495b-9370-853cd873bb07" {return "MicrosoftCC"}
   "f4f024c8-d611-4748-a7e0-02b6e754c0fe" {return "MicrosoftDC"}
   "b61285dd-d9f7-41f2-9757-8f61cba4e9c8" {return "MicrosoftDevMain"}
   "9a3b7ff2-58ed-40fd-add5-1e5158059d1c" {return "MicrosoftFRDC"}
   "1d2d2ea6-1680-4c56-ac58-a441c8c24ff9" {return "MicrosoftLTSC"}
   "86752282-5841-4120-ac80-db03ae6b5fdb" {return "MicrosoftLTSC2021"}
   "C02D8FE6-5242-4DA8-972F-82EE55E00671" {return "MicrosoftLTSC2024"}
   "2e148de9-61c8-4051-b103-4af54baffbb4" {return "InsidersLTSC"}
   "12f4f6ad-fdea-4d2a-a90f-17496cc19a48" {return "InsidersLTSC2021"}
   "20481F5C-C268-4624-936C-52EB39DDBD97" {return "InsidersLTSC2024"}
   "0002c1ba-b76b-4af9-b1ee-ae2ad587371f" {return "InsidersMEC"}
  }
  return "Null"
}
Function Get-Office-Apps {
   param (
     [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName)]
     [string] $FFN
   )
  
  $ProgressPreference = 'SilentlyContinue'    # Subsequent calls do not display UI.
  $URI = 'https://clients.config.office.net/releases/v1.0/OfficeReleases'
  $URI = 'https://mrodevicemgr.officeapps.live.com/mrodevicemgrsvc/api/v2/C2RReleaseData'
  $REQ = IWR $URI -ea 0

  if (-not $REQ) {
    return $null
  }

  $Json = $REQ.Content | ConvertFrom-Json
  $Json|Sort-Object FFN|select @{Name='Channel'; Expr={$_.FFN|Convert-To-Channel}},FFN,@{Name='Build'; Expr={$_.AvailableBuild}},@{Name='System'; Expr={$_.Name|Convert-To-System}}
}
Function Office_Offline_Install (
  [Parameter(ValueFromPipeline, ValueFromPipelineByPropertyName, Mandatory)]
  [String] $Location)
{
  
Function Get-Lang {
  $oLang = @{}
  $oLang.Add(1033,"English")
  $oLang.Add(1078,"Afrikaans")
  $oLang.Add(1052,"Albanian")
  $oLang.Add(1118,"Amharic")
  $oLang.Add(1025,"Arabic")
  $oLang.Add(1067,"Armenian")
  $oLang.Add(1101,"Assamese")
  $oLang.Add(1068,"Azerbaijani Latin")
  $oLang.Add(2117,"Bangla Bangladesh")
  $oLang.Add(1093,"Bangla Bengali India")
  $oLang.Add(1069,"Basque Basque")
  $oLang.Add(1059,"Belarusian")
  $oLang.Add(5146,"Bosnian")
  $oLang.Add(1026,"Bulgarian")
  $oLang.Add(2051,"Catalan Valencia")
  $oLang.Add(1027,"Catalan")
  $oLang.Add(2052,"Chinese Simplified")
  $oLang.Add(1028,"Chinese Traditional")
  $oLang.Add(1050,"Croatian")
  $oLang.Add(1029,"Czech")
  $oLang.Add(1030,"Danish")
  $oLang.Add(1164,"Dari")
  $oLang.Add(1043,"Dutch")
  $oLang.Add(2057,"English UK")
  $oLang.Add(1061,"Estonian")
  $oLang.Add(1124,"Filipino")
  $oLang.Add(1035,"Finnish")
  $oLang.Add(3084,"French Canada")
  $oLang.Add(1036,"French")
  $oLang.Add(1110,"Galician")
  $oLang.Add(1079,"Georgian")
  $oLang.Add(1031,"German")
  $oLang.Add(1032,"Greek")
  $oLang.Add(1095,"Gujarati")
  $oLang.Add(1128,"Hausa Nigeria")
  $oLang.Add(1037,"Hebrew")
  $oLang.Add(1081,"Hindi")
  $oLang.Add(1038,"Hungarian")
  $oLang.Add(1039,"Icelandic")
  $oLang.Add(1136,"Igbo")
  $oLang.Add(1057,"Indonesian")
  $oLang.Add(2108,"Irish")
  $oLang.Add(1040,"Italian")
  $oLang.Add(1041,"Japanese")
  $oLang.Add(1099,"Kannada")
  $oLang.Add(1087,"Kazakh")
  $oLang.Add(1107,"Khmer")
  $oLang.Add(1089,"KiSwahili")
  $oLang.Add(1159,"Kinyarwanda")
  $oLang.Add(1111,"Konkani")
  $oLang.Add(1042,"Korean")
  $oLang.Add(1088,"Kyrgyz")
  $oLang.Add(1062,"Latvian")
  $oLang.Add(1063,"Lithuanian")
  $oLang.Add(1134,"Luxembourgish")
  $oLang.Add(1071,"Macedonian")
  $oLang.Add(1086,"Malay Latin")
  $oLang.Add(1100,"Malayalam")
  $oLang.Add(1082,"Maltese")
  $oLang.Add(1153,"Maori")
  $oLang.Add(1102,"Marathi")
  $oLang.Add(1104,"Mongolian")
  $oLang.Add(1121,"Nepali")
  $oLang.Add(2068,"Norwedian Nynorsk")
  $oLang.Add(1044,"Norwegian Bokmal")
  $oLang.Add(1096,"Odia")
  $oLang.Add(1123,"Pashto")
  $oLang.Add(1065,"Persian")
  $oLang.Add(1045,"Polish")
  $oLang.Add(1046,"Portuguese Brazilian")
  $oLang.Add(2070,"Portuguese Portugal")
  $oLang.Add(1094,"Punjabi Gurmukhi")
  $oLang.Add(3179,"Quechua")
  $oLang.Add(1048,"Romanian")
  $oLang.Add(1047,"Romansh")
  $oLang.Add(1049,"Russian")
  $oLang.Add(1074,"Setswana")
  $oLang.Add(1169,"Scottish Gaelic")
  $oLang.Add(7194,"Serbian Bosnia")
  $oLang.Add(10266,"Serbian Serbia")
  $oLang.Add(9242,"Serbian")
  $oLang.Add(1132,"Sesotho sa Leboa")
  $oLang.Add(2137,"Sindhi Arabic")
  $oLang.Add(1115,"Sinhala")
  $oLang.Add(1051,"Slovak")
  $oLang.Add(1060,"Slovenian")
  $oLang.Add(3082,"Spanish")
  $oLang.Add(2058,"Spanish Mexico")
  $oLang.Add(1053,"Swedish")
  $oLang.Add(1097,"Tamil")
  $oLang.Add(1092,"Tatar Cyrillic")
  $oLang.Add(1098,"Telugu")
  $oLang.Add(1054,"Thai")
  $oLang.Add(1055,"Turkish")
  $oLang.Add(1090,"Turkmen")
  $oLang.Add(1058,"Ukrainian")
  $oLang.Add(1056,"Urdu")
  $oLang.Add(1152,"Uyghur")
  $oLang.Add(1091,"Uzbek")
  $oLang.Add(1066,"Vietnamese")
  $oLang.Add(1106,"Welsh")
  $oLang.Add(1160,"Wolof")
  $oLang.Add(1130,"Yoruba")
  $oLang.Add(1076,"isiXhosa")
  $oLang.Add(1077,"isiZulu")
  return $oLang
}
Function Get-Culture {
  $oLang = @{}
  $oLang.Add(1033,"en-us")
  $oLang.Add(1078,"af-za")
  $oLang.Add(1052,"sq-al")
  $oLang.Add(1118,"am-et")
  $oLang.Add(1025,"ar-sa")
  $oLang.Add(1067,"hy-am")
  $oLang.Add(1101,"as-in")
  $oLang.Add(1068,"az-latn-az")
  $oLang.Add(2117,"bn-bd")
  $oLang.Add(1093,"bn-in")
  $oLang.Add(1069,"eu-es")
  $oLang.Add(1059,"be-by")
  $oLang.Add(5146,"bs-latn-ba")
  $oLang.Add(1026,"bg-bg")
  $oLang.Add(2051,"ca-es-valencia")
  $oLang.Add(1027,"ca-es")
  $oLang.Add(2052,"zh-cn")
  $oLang.Add(1028,"zh-tw")
  $oLang.Add(1050,"hr-hr")
  $oLang.Add(1029,"cs-cz")
  $oLang.Add(1030,"da-dk")
  $oLang.Add(1164,"prs-af")
  $oLang.Add(1043,"nl-nl")
  $oLang.Add(2057,"en-GB")
  $oLang.Add(1061,"et-ee")
  $oLang.Add(1124,"fil-ph")
  $oLang.Add(1035,"fi-fi")
  $oLang.Add(3084,"fr-CA")
  $oLang.Add(1036,"fr-fr")
  $oLang.Add(1110,"gl-es")
  $oLang.Add(1079,"ka-ge")
  $oLang.Add(1031,"de-de")
  $oLang.Add(1032,"el-gr")
  $oLang.Add(1095,"gu-in")
  $oLang.Add(1128,"ha-Latn-NG")
  $oLang.Add(1037,"he-il")
  $oLang.Add(1081,"hi-in")
  $oLang.Add(1038,"hu-hu")
  $oLang.Add(1039,"is-is")
  $oLang.Add(1136,"ig-NG")
  $oLang.Add(1057,"id-id")
  $oLang.Add(2108,"ga-ie")
  $oLang.Add(1040,"it-it")
  $oLang.Add(1041,"ja-jp")
  $oLang.Add(1099,"kn-in")
  $oLang.Add(1087,"kk-kz")
  $oLang.Add(1107,"km-kh")
  $oLang.Add(1089,"sw-ke")
  $oLang.Add(1159,"rw-RW")
  $oLang.Add(1111,"kok-in")
  $oLang.Add(1042,"ko-kr")
  $oLang.Add(1088,"ky-kg")
  $oLang.Add(1062,"lv-lv")
  $oLang.Add(1063,"lt-lt")
  $oLang.Add(1134,"lb-lu")
  $oLang.Add(1071,"mk-mk")
  $oLang.Add(1086,"ms-my")
  $oLang.Add(1100,"ml-in")
  $oLang.Add(1082,"mt-mt")
  $oLang.Add(1153,"mi-nz")
  $oLang.Add(1102,"mr-in")
  $oLang.Add(1104,"mn-mn")
  $oLang.Add(1121,"ne-np")
  $oLang.Add(2068,"nn-no")
  $oLang.Add(1044,"nb-no")
  $oLang.Add(1096,"or-in")
  $oLang.Add(1123,"ps-AF")
  $oLang.Add(1065,"fa-ir")
  $oLang.Add(1045,"pl-pl")
  $oLang.Add(1046,"pt-br")
  $oLang.Add(2070,"pt-pt")
  $oLang.Add(1094,"pa-in")
  $oLang.Add(3179,"quz-pe")
  $oLang.Add(1048,"ro-ro")
  $oLang.Add(1047,"rm-CH")
  $oLang.Add(1049,"ru-ru")
  $oLang.Add(1074,"tn-ZA")
  $oLang.Add(1169,"gd-gb")
  $oLang.Add(7194,"sr-cyrl-ba")
  $oLang.Add(10266,"sr-cyrl-rs")
  $oLang.Add(9242,"sr-latn-rs")
  $oLang.Add(1132,"nso-ZA")
  $oLang.Add(2137,"sd-arab-pk")
  $oLang.Add(1115,"si-lk")
  $oLang.Add(1051,"sk-sk")
  $oLang.Add(1060,"sl-si")
  $oLang.Add(3082,"es-es")
  $oLang.Add(2058,"es-MX")
  $oLang.Add(1053,"sv-se")
  $oLang.Add(1097,"ta-in")
  $oLang.Add(1092,"tt-ru")
  $oLang.Add(1098,"te-in")
  $oLang.Add(1054,"th-th")
  $oLang.Add(1055,"tr-tr")
  $oLang.Add(1090,"tk-tm")
  $oLang.Add(1058,"uk-ua")
  $oLang.Add(1056,"ur-pk")
  $oLang.Add(1152,"ug-cn")
  $oLang.Add(1091,"uz-latn-uz")
  $oLang.Add(1066,"vi-vn")
  $oLang.Add(1106,"cy-gb")
  $oLang.Add(1160,"wo-SN")
  $oLang.Add(1130,"yo-NG")
  $oLang.Add(1076,"xh-ZA")
  $oLang.Add(1077,"zu-ZA")
  return $oLang
}
Function Get-Channels {
  $oProd = @{}
  $oProd.Add("BetaChannel","5440fd1f-7ecb-4221-8110-145efaa6372f")
  $oProd.Add("Current","492350f6-3a01-4f97-b9c0-c7c6ddf67d60")
  $oProd.Add("CurrentPreview","64256afe-f5d9-4f86-8936-8840a6a4f5be")
  $oProd.Add("DogfoodCC","f3260cf1-a92c-4c75-b02e-d64c0a86a968")
  $oProd.Add("DogfoodDCEXT","c4a7726f-06ea-48e2-a13a-9d78849eb706")
  $oProd.Add("DogfoodDevMain","ea4a4090-de26-49d7-93c1-91bff9e53fc3")
  $oProd.Add("DogfoodFRDC","834504cc-dc55-4c6d-9e71-e024d0253f6d")
  $oProd.Add("InsidersLTSC","2e148de9-61c8-4051-b103-4af54baffbb4")
  $oProd.Add("InsidersLTSC2021","12f4f6ad-fdea-4d2a-a90f-17496cc19a48")
  $oProd.Add("InsidersLTSC2024","20481F5C-C268-4624-936C-52EB39DDBD97")
  $oProd.Add("InsidersMEC","0002c1ba-b76b-4af9-b1ee-ae2ad587371f")   
  $oProd.Add("MicrosoftCC","5462eee5-1e97-495b-9370-853cd873bb07")
  $oProd.Add("MicrosoftDC","f4f024c8-d611-4748-a7e0-02b6e754c0fe")
  $oProd.Add("MicrosoftDevMain","b61285dd-d9f7-41f2-9757-8f61cba4e9c8")
  $oProd.Add("MicrosoftFRDC","9a3b7ff2-58ed-40fd-add5-1e5158059d1c")
  $oProd.Add("MicrosoftLTSC","1d2d2ea6-1680-4c56-ac58-a441c8c24ff9")
  $oProd.Add("MicrosoftLTSC2021","86752282-5841-4120-ac80-db03ae6b5fdb")
  $oProd.Add("MicrosoftLTSC2024","C02D8FE6-5242-4DA8-972F-82EE55E00671")
  $oProd.Add("MonthlyEnterprise","55336b82-a18d-4dd6-b5f6-9e5095c314a6")
  $oProd.Add("PerpetualVL2019","f2e724c1-748f-4b47-8fb8-8e0d210e9208")
  $oProd.Add("PerpetualVL2021","5030841d-c919-4594-8d2d-84ae4f96e58e")
  $oProd.Add("PerpetualVL2024","7983BAC0-E531-40CF-BE00-FD24FE66619C")
  $oProd.Add("SemiAnnual","7ffbc6bf-bc32-4f92-8982-f9dd17fd3114")
  $oProd.Add("SemiAnnualPreview","b8f9b850-328d-4355-9145-c59439a0c4cf")
  return $oProd
}

  $IsX32=$Null
  $IsX64=$Null
  $file = $Null

  $IgnoreCase = [Text.RegularExpressions.RegexOptions]::IgnoreCase
  if (($Location -eq $Null) -or ((Test-Path $Location) -eq $false)) {
    $Location
    throw "ERROR: BAD PATH"
  }

  Set-Location $Location
  $Fixed = Get-Location

  if ((-not(Test-Path "$Fixed\Office\Data\v64.cab")) -and
    (-not(Test-Path "$Fixed\Office\Data\v32.cab"))) {
    throw "ERROR: MISSING CAB FILES"
  }

  if (Test-Path "$Fixed\Office\Data\v32.cab") {
    $IsX32=$true
  }
  if (Test-Path "$Fixed\Office\Data\v64.cab") {
    $IsX64=$true
  }

  Switch ([intptr]::Size) {
    4 { 
        if (-not (
          $IsX32)) {
            throw "ERROR: SYSTEM NOT MATCH"  }
        $file = "$Fixed\Office\Data\v32.cab"
        $IsX64 = $Null
      }
    8 {
        if ($IsX64 -and (
          Test-Path "$Fixed\Office\Data\v64.cab")) {
            $IsX32 = $Null
            $file = "$Fixed\Office\Data\v64.cab" }
	    if ($IsX32 -and (
		  Test-Path "$Fixed\Office\Data\v32.cab")) {
            $IsX64 = $Null
            $file = "$Fixed\Office\Data\v32.cab" }
      }
  }

  ri @(Join-Path $env:TEMP VersionDescriptor.xml) -Force -ea 0
  Expand $file -f:VersionDescriptor.xml $env:TEMP *>$Null
  if (!(Test-path(
    @(Join-Path $env:TEMP VersionDescriptor.xml)))) {
      throw "ERROR: FAIL EXTRACT XML FILE"
  }

  $oXml = Get-Content @(Join-Path $env:TEMP VersionDescriptor.xml) -ea 0
  if (!$oXml) {
    throw "ERROR: FAIL READ XML FILE"
  }
  
  # find FFNRoot value
  $DeliveryMechanism = $oXml | ? {$_ -match 'DeliveryMechanism FFNRoot'} | select -First 1
  $FFNRoot = $DeliveryMechanism.Substring(30,36)
  $sUrl="http://officecdn.microsoft.com/pr/$FFNRoot"

  $oLang = Get-Lang
  $oData = GCI "$(Get-Location)\Office\Data"
  $oVer = $oData[0].Name
  $oData = GCI "$(Get-Location)\Office\Data\$oVer"
  $fPatX64 = '^(sp64)(.*)(\.cab.*)$'
  $fPatX86 = '^(sp32)(.*)(\.cab.*)$'
  $oList = $oData | ? {
    ($IsX32 -and ([REGEX]::IsMatch($_.Name,$fPatX86,$IgnoreCase))) -or (
      ($IsX64 -and ([REGEX]::IsMatch($_.Name,$fPatX64,$IgnoreCase))))
  } | %{$_.Name} | Sort
 
  $oIDS = $oList.Substring(4,4)
  $sLang = $oIDS | % {  $oLang[[INT]$_]} | OGV -Title "Select Language" -PassThru
  if (-not $sLang) {
    return;
  }

  $rPat = '^(.*)(ProductReleaseId Name=)(.*)>$'
  $oApps = $oXml|?{[REGEX]::IsMatch($_,$rPat,$IgnoreCase)}|%{$_.SubString(28,$_.Length-28-2)}|sort|OGV -title "Found Office Apps" -PassThru
  if (-not $oApps) {
    return;
  }

  $mLang = $sLang|%{
    $oLang.GetEnumerator()|? Value -EQ $_
  }
  
  # Start set values
  $type = "LOCAL"
  $bUrl = """$($Fixed)"""
  $misc = "flt.useoutlookshareaddon=unknown flt.useofficehelperaddon=unknown"

  $culture = ''
  $oCul = Get-Culture

  $mLang | % {$culture += $oCul[[INT]$_.Name]+'_'}
  $sCulture = $culture.SubString(0,$culture.Length-1)
  $mCulture = $culture.Split('_')[0]

  $AppList = ''
  $oApps | % {$AppList += "$($_).16_$($sCulture)_x-none|"}
  $sAppList = $AppList.TrimEnd('|')
  
  gwmi Win32_Service | ? Name -Match "WSearch|ClickToRunSvc" | % {
    $_.StopService() | Out-Null }
  
  Get-Service -Name @("WSearch", "ClickToRunSvc") | 
    Stop-Service -Force -PassThru | Out-Null
  
  if ($IsX32) {
	$vSys = "x86"
    MD "$ENV:ProgramFiles(x86)\Common Files\Microsoft Shared\ClickToRun" -ea 0 | Out-Null
    Expand "$(Get-Location)\Office\Data\$oVer\i320.cab" -f:* "$ENV:ProgramFiles(x86)\Common Files\Microsoft Shared\ClickToRun" *>$Null
    Push-Location "$ENV:ProgramFiles(x86)\Common Files\Microsoft Shared\ClickToRun\"
  }

  if ($IsX64) {
	$vSys = "x64"
    MD "$ENV:ProgramFiles\Common Files\Microsoft Shared\ClickToRun" -ea 0 | Out-Null
    Expand "$(Get-Location)\Office\Data\$oVer\i640.cab" -f:* "$ENV:ProgramFiles\Common Files\Microsoft Shared\ClickToRun" *>$Null
    Push-Location "$ENV:ProgramFiles\Common Files\Microsoft Shared\ClickToRun\"
  }
  
  start OfficeClickToRun.exe -args "platform=$vSys culture=$mCulture productstoadd=$sAppList cdnbaseurl.16=$sUrl baseurl.16=$bUrl version.16=$oVer mediatype.16=$type sourcetype.16=$type updatesenabled.16=True acceptalleulas.16=True displaylevel=True bitnessmigration=False deliverymechanism=$FFNRoot $misc"
}

if (-not($ARGS)) {
  Office_Offline_Install (Get-Location).Path
  return
}

Office_Offline_Install $ARGS[0]
return
# SIG # Begin signature block
# MIIFoQYJKoZIhvcNAQcCoIIFkjCCBY4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQU0rSZpN0atjbPX8rEmX3PODuX
# 9MCgggM2MIIDMjCCAhqgAwIBAgIQSe49BypwLKhOHkzMeRKLBzANBgkqhkiG9w0B
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
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFCNw
# 0F1Snqnm4h8SDXVQae4J3+mxMA0GCSqGSIb3DQEBAQUABIIBALepD4z9shi+KPsQ
# aHfxQvHlH3gbs+LEgmZ86MrB/yioTT3t8j6eOlYQ2BRYN7Y9WCYjLidC1wmepNua
# UxyqHiyiQYD6UnxeWRoA2lu5twAKoEeWkIH38BgkTNnPtKrRpd7d+vnLuI7S9I9q
# 4SQg0kTfYejxXg0m+9HcfJR8RkkWGUrnYB6/6t+ertmM1eErlvrq22Xnf8ZY51+N
# ZdJLaH9Hc65gAEXwM+/jtrH1LWkRgTxSR1E4eYpOFS00JxPsdSnlnfiWSKwyXcHp
# hjOOiY4ekOpFUiFK+u8QX0wCpMgdq5nqx+tQrFbvUT3unq1WFD4f3dTDEqEukyDy
# LGlRPe0=
# SIG # End signature block
