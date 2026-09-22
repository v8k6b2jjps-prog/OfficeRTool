cls
Write-Host

try {
  Set-Location "$env:ProgramFiles\Microsoft Office\root\Licenses16" }
catch {
  return }

$LicensingService = gwmi SoftwareLicensingService -ea 1
if (-not $LicensingService) {
  return }

$file_list = dir * -Name
$loc = (Get-Location).Path
$Selection = $file_list | ogv -Title "License installer - Helper" -OutputMode Multiple
if ($Selection) {
    
    $file_list | ? {$_ -like "pkeyconfig*" -or $_ -like "Client*"} | % {
      $LicenseFile = Join-Path $loc $_
      Write-Host "Install License: $($LicenseFile)"
      $LicensingService.InstallLicense(
        [IO.FILE]::ReadAllText($LicenseFile)) | Out-Null}

    $Selection | % {
      $LicenseFile = Join-Path $loc $_
      Write-Host "Install License: $($LicenseFile)"
      $LicensingService.InstallLicense(
        [IO.FILE]::ReadAllText($LicenseFile)) | Out-Null
    }
}

Start-Sleep 4
return
# SIG # Begin signature block
# MIIFoQYJKoZIhvcNAQcCoIIFkjCCBY4CAQExCzAJBgUrDgMCGgUAMGkGCisGAQQB
# gjcCAQSgWzBZMDQGCisGAQQBgjcCAR4wJgIDAQAABBAfzDtgWUsITrck0sYpfvNR
# AgEAAgEAAgEAAgEAAgEAMCEwCQYFKw4DAhoFAAQUrM9BJNywONcjKO2/e40shH/y
# BPigggM2MIIDMjCCAhqgAwIBAgIQSe49BypwLKhOHkzMeRKLBzANBgkqhkiG9w0B
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
# HAYKKwYBBAGCNwIBCzEOMAwGCisGAQQBgjcCARUwIwYJKoZIhvcNAQkEMRYEFAVS
# 2Co1EvMlEjFyG6EwWC4MAf4IMA0GCSqGSIb3DQEBAQUABIIBAK16EgG51ri8KIUD
# gz3gHH6JG2Zcc5QOYyDFiIvYVV/h2DPsm99ilNn0tECSmpVUoxO1T2BVZNrRCaqE
# hrURLgen0lkQxhTLLqgqaRxtfI3uglzDo89o4/+Xyh71t0N27AZA5xhU19VqoDMt
# a5zQLKTwvNWUsTUMONnJ/zltb1qlFr6cSGQOGoY6uEddnuUel6SJZW9ZtB9m3LZV
# Va2jddjTYD90pgQyocREMvl5cnwKvk1v2Z5rSHGB4nq6+yiYuqPWIIG4CaeESLMM
# qQ6FjLa3umuH7+nPpWVEHfh1Bp4tau6DCT+cG1mGeIdBiWkQHQlYFjmEbBZpi31T
# pemlWws=
# SIG # End signature block
