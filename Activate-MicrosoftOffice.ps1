function Activate-MicrosoftOffice {
#.SYNOPSIS
# Activate Free Microsoft Office Professional Plus 2016 - 2021
# ARBITRARY VERSION NUMBER:  1.1.0
# AUTHOR:  Tyler McCann (@tylerdotrar)
#
#.DESCRIPTION
# Automatically detect and activate locally installed Microsoft Office with a Professional Plus
# KMS client key using a publically available KMS server.  This script supports both 32-bit and
# 64-bit versions of Microsoft Office 2016, 2019, and 2021.  Must run elevated.
#
# Parameters:
#   -KMSserver    -->  Domain/IP of specified KMS server
#   -KMSport      -->  Port of specified KMS server
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

    
    Param ( 
        [string]$KMSserver,
        [int]$KMSport
    )


    # Determine if user has elevated privileges
    $User    = [Security.Principal.WindowsIdentity]::GetCurrent();
    $isAdmin = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if (!$isAdmin) { return (Write-Host 'This function requires elevated privileges.' -ForegroundColor Red) }
    

    ## Step 1
    Write-Host 'Determining if Office 64-bit or 32-bit version is installed...' -ForegroundColor Yellow
    $PreActivate  = "$PWD"
    $OfficeDir    = "$env:ProgramFiles\Microsoft Office\Office16"
    $OfficeDirx86 = "${env:ProgramFiles(x86)}\Microsoft Office\Office16"

    if (Test-Path -LiteralPath $OfficeDir)        { cd -LiteralPath $OfficeDir    ; Write-Host "- 64-bit`n" }
    elseif (Test-Path -LiteralPath $OfficeDirx86) { cd -LiteralPath $OfficeDirx86 ; Write-host "- 32-bit`n" }
    else { return (Write-Host "- Failed to find Microsoft Office.`n" -ForegroundColor Red) }

    
    ## Step 2
    Write-Host 'Determining version of Office to activate...' -ForegroundColor Yellow
    $OfficeStatus  = cscript /nologo ospp.vbs /dstatus
    $OfficeVersion = ($OfficeStatus | Select-String -Pattern "ProPlus(\d{4})?").Matches.Value
    $CurrentKey    = (($OfficeStatus | Select-String -Pattern "product key:") -as [string]).Split(' ')[-1]

    if (!$OfficeVersion) {
        cd -LiteralPath $PreActivate
        return (Write-Host "- Initial Office license not detected; please first-time open a Microsoft Office executable or reboot your computer.`n" -ForegroundColor Red)
    }
    
    # Microsoft Office KMS Client Key Table
    $KeyTable = @(
        [PSCustomObject]@{OfficeName="Microsoft Office 2021";OfficeVersion="ProPlus2021";Key="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"},
        [PSCustomObject]@{OfficeName="Microsoft Office 2019";OfficeVersion="ProPlus2019";Key="NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"},
        [PSCustomObject]@{OfficeName="Microsoft Office 2016";OfficeVersion="ProPlus"    ;Key="XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"}
    )

    $KeyTable | % { if ($OfficeVersion -eq $_.OfficeVersion) { $ActivationKey = $_.Key ; $OfficeName = $_.OfficeName } }
    Write-Host "- $OfficeName`n"


    ## Step 3
    if ($KMSserver) { Write-Host 'Testing connection to input KMS server for activation...'     -ForegroundColor Yellow }
    else            { Write-Host 'Determining which public KMS server to use for activation...' -ForegroundColor Yellow }

    # User input KMS server
    if ($KMSserver) { $ServerList = @("$KMSserver") }
    else {
        $ServerList += @(
            'e8.us.to',
            'e9.us.to',
            'kms8.msguides.com',
            'kms9.msguides.com'
        )
    }

    foreach ($Server in $ServerList) {
        # User input KMS port
        if (!$KMSport) { $KMsport = 1688 }
        $Connection = Test-NetConnection -ComputerName $Server -Port $KMSport

        if ($Connection.TcpTestSucceeded) { $KMS_Server = $Server ; break }
    }

    if (!$KMS_Server) {
        cd -LiteralPath $PreActivate
        return (Write-Host "- Failed to connect to any KMS servers; please try a different server.`n(Use 'Activate-MicrosoftOffice' with params '-KMSserver' & '-KMSport').`n" -ForegroundColor Red)
    }
    Write-Host "- ${KMS_Server}`n"


    ## Step 4
    Write-Host 'Converting retail license to volume license...' -ForegroundColor Yellow
    $Licenses = (Get-ChildItem "..\root\Licenses16\${OfficeVersion}VL_KMS*.xrm-ms").FullName 2>$NULL

    if ($Licenses -eq $NULL) { Write-Host "- Skipping`n" }
    else { $Licenses | % { cscript /nologo ospp.vbs /inslic:$_ } }


    ## Step 5
    Write-Host 'Activating Microsoft Office using KMS client key...' -ForegroundColor Yellow
    cscript /nologo ospp.vbs /sethst:${KMS_Server}
    cscript /nologo ospp.vbs /setprt:$KMSport
    cscript /nologo ospp.vbs /unpkey:$CurrentKey
    cscript /nologo ospp.vbs /inpkey:$ActivationKey
    cscript /nologo ospp.vbs /act

    # Return to original directory post-activation
    cd -LiteralPath $PreActivate
    return (Write-Host 'Complete.' -ForegroundColor Green)
}