function Activate-MicrosoftOffice {
#.SYNOPSIS
# Activate Microsoft Office Professional Plus 2016 - 2021
# ARBITRARY VERSION NUMBER:  1.0.2
# AUTHOR:  Tyler McCann (@tylerdotrar)
#
#.DESCRIPTION
# Automatically detect and activate locally installed Microsoft Office with a Professional Plus
# KMS client key using a publically available KMS server.  This script supports both 32-bit and
# 64-bit versions of Microsoft Office 2016, 2019, and 2021.  Must run elevated.
#
#
# Official Microsoft Office Downloads:
#
#  - Microsoft Office Professional Plus 2021:
#  - Download URL : https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img
#
#  - Microsoft Office Professional Plus 2019:
#  - Download URL : https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img
#
#  - Microsoft Office Professional Plus 2016:
#  - Download URL : <N/A>
#
#.LINK
# https://github.com/tylerdotrar/Activate-MicrosoftOffice


    # Determine if user has elevated privileges
    $User    = [Security.Principal.WindowsIdentity]::GetCurrent();
    $isAdmin = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if (!$isAdmin) { return (Write-Host 'This function requires elevated privileges.' -ForegroundColor Red) }
    

    ## Step 1
    Write-Host 'Determining if Office 64-bit or 32-bit version is installed...' -ForegroundColor Yellow
    $PreActivate  = "$PWD"
    $OfficeDir    = "$env:ProgramFiles\Microsoft Office\Office16"
    $OfficeDirx86 = "${env:ProgramFiles(x86)}\Microsoft Office\Office16"

    if (Test-Path -LiteralPath $OfficeDir)        { cd -LiteralPath $OfficeDir    ; Write-Host ' - 64-bit' }
    elseif (Test-Path -LiteralPath $OfficeDirx86) { cd -LiteralPath $OfficeDirx86 ; Write-host ' - 32-bit' }
    else { return (Write-Host ' - Failed to find Microsoft Office.  Exiting script.' -ForegroundColor Red) }

    
    ## Step 2
    Write-Host 'Determining version of Office to activate...' -ForegroundColor Yellow
    $OfficeStatus  = cscript /nologo ospp.vbs /dstatus
    $OfficeVersion = ($OfficeStatus | Select-String -Pattern "ProPlus(\d{4})?").Matches.Value
    $CurrentKey    = (($OfficeStatus | Select-String -Pattern "product key:") -as [string]).Split(' ')[-1]
    
    # Microsoft Office KMS Client Key Table
    $KeyTable = @(
        [PSCustomObject]@{OfficeName="Microsoft Office 2021";OfficeVersion="ProPlus2021";Key="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"},
        [PSCustomObject]@{OfficeName="Microsoft Office 2019";OfficeVersion="ProPlus2019";Key="NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"},
        [PSCustomObject]@{OfficeName="Microsoft Office 2016";OfficeVersion="ProPlus"    ;Key="XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"}
    )

    $KeyTable | % { if ($OfficeVersion -eq $_.OfficeVersion) { $ActivationKey = $_.Key ; $OfficeName = $_.OfficeName } }
    Write-Host " - $OfficeName"


    ## Step 3
    Write-Host 'Determining which public KMS server to use for activation...' -ForegroundColor Yellow
    $ServerList = @(
        'e8.us.to',
        'e9.us.to',
        'kms8.msguides.com',
        'kms9.msguides.com'
    )

    foreach ($Server in $ServerList) {
        $Connection = Test-NetConnection -ComputerName $Server -Port 1688
        if ($Connection.TcpTestSucceeded) { $KMS_Server = $Server ; break }
    }

    if (!$KMS_Server) { return (Write-Host ' - Failed to connect to any KMS servers. Exiting script.' -ForegroundColor Red) }
    Write-Host " - ${KMS_Server}"


    ## Step 4
    Write-Host 'Converting retail license to volume license...' -ForegroundColor Yellow
    $Licenses = (Get-ChildItem "..\root\Licenses16\${OfficeVersion}VL_KMS*.xrm-ms").FullName

    if ($Licenses -eq $NULL) { Write-Host ' - Skipping' }
    else { $Licenses | % { cscript /nologo ospp.vbs /inslic:$_ } }


    ## Step 5
    Write-Host 'Activating Microsoft Office using KMS client key...' -ForegroundColor Yellow
    cscript /nologo ospp.vbs /sethst:${KMS_Server}
    cscript /nologo ospp.vbs /setprt:1688
    cscript /nologo ospp.vbs /unpkey:$CurrentKey
    cscript /nologo ospp.vbs /inpkey:$ActivationKey
    cscript /nologo ospp.vbs /act

    # Return to original directory post-activation
    cd -LiteralPath $PreActivate
    return (Write-Host 'Complete.' -ForegroundColor Green)
}