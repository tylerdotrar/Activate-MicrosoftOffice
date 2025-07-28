function Activate-MicrosoftOffice {
#.SYNOPSIS
# Activate Microsoft Office 2016 - 2024 via free Professional Plus KMS Client Keys
# Arbitrary Version Number: 2.0.3
# Author:  Tyler McCann (@tylerdotrar)
#
#.DESCRIPTION
# Automatically detect and activate a locally installed Microsoft Office instance with a Professional Plus KMS client key,
# using either a publicly available KMS server (default) or a specified KMS server.  This script supports both 32-bit and
# 64-bit versions of Microsoft Office 2016, 2019, 2021, and 2024.
#
# (Note: Microsoft Office 2024 requires the '-Office2024' parameter due to being unable to detect the default license.)
#
# This script will set a registry key to disable automatic updates and (when using Microsoft Office 2016 or 2019) will also
# rollback the Office version to '16.0.13801.20266' to prevent the annoying 'GET GENUINE OFFICE' banner that plagued earlier
# versions of this script when using KMS clients keys with public KMS servers.  To disable this functionality, run this
# script with the '-DontRollback' parameter.
#
# Notes:
# o This script must run with elevated privileges (i.e., as Administrator).
# o This script contains a list of known public KMS servers and will automatically attempt to use
#   them if the user doesn't specify a target KMS server ('-KMSserver') and/or port ('-KMSport').
#
# Parameters:
#   -KMSserver     -->  Target domain/IP of a specified KMS server.
#   -KMSport       -->  Target port of a specified KMS server.
#   -Office2024    -->  Activate Microsoft Office 2024 (due to license auto-detection failing).
#   -DontRollback  -->  Do not rollback Office version and leave automatic updates enabled.
#   -Help          -->  Return Get-Help information.
#
# Official Microsoft Office Downloads:
#   > Office 2024 (ProPlus) : https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2024Retail.img
#   > Office 2021 (ProPlus) : https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img
#   > Office 2019 (ProPlus) : https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img
#   > Office 2016 (ProPlus) : https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-us/ProPlusRetail.img
#
#.LINK
# https://github.com/tylerdotrar/Activate-MicrosoftOffice

    
    Param ( 
        [string]$KMSserver,
        [int]   $KMSport,
        [switch]$Office2024,
        [switch]$DontRollback,
        [switch]$Help
    )


    # Return Get-Help Information
    if ($Help) { return (Get-Help Activate-MicrosoftOffice) }


    # Determine if user has elevated privileges
    $User    = [Security.Principal.WindowsIdentity]::GetCurrent();
    $isAdmin = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if (!$isAdmin) { return (Write-Host "[!] This function requires elevated privileges.`n" -ForegroundColor Red) }
    

    ## Step 1: Verify Office is installed
    Write-Host '[+] Determining if Office 64-bit or 32-bit version is installed...' -ForegroundColor Yellow
    $PreActivate  = "$PWD"
    $OfficeDir    = "${env:ProgramFiles}\Microsoft Office\Office16"
    $OfficeDirx86 = "${env:ProgramFiles(x86)}\Microsoft Office\Office16"

    if (Test-Path -LiteralPath $OfficeDir)        { cd -LiteralPath $OfficeDir    ; Write-Host " o  64-bit`n" }
    elseif (Test-Path -LiteralPath $OfficeDirx86) { cd -LiteralPath $OfficeDirx86 ; Write-host " o  32-bit`n" }
    else { return (Write-Host "[!] Failed to find Microsoft Office.`n" -ForegroundColor Red) }

    
    ## Step 2: Determine version of Office installed
    Write-Host "[+] Determining version of Office to activate..." -ForegroundColor Yellow
    $OfficeStatus  = cscript /nologo ospp.vbs /dstatus
    $OfficeVersion = ($OfficeStatus | Select-String -Pattern "ProPlus(\d{4})?").Matches.Value
    $CurrentKey    = (($OfficeStatus | Select-String -Pattern "product key:") -as [string]).Split(' ')[-1]

    if (!$OfficeVersion) {
        
        if ($Office2024) {
            Write-Host " o  No license detected; continuing script for Microsoft Office 2024."
            $OfficeVersion = 'ProPlus2024'
        }
        else {
            cd -LiteralPath $PreActivate
            Write-Host "[!] Initial Office license not detected." -ForegroundColor Red
            Write-host " o  Please first-time open a Microsoft Office executable and/or reboot your computer." -ForegroundColor Red
            Write-Host " o  If using Microsoft Office 2024, please use the '-Office2024' parameter.`n" -ForegroundColor Red
            
            return
        }
    }
    
    # Microsoft Office KMS Client Key Table
    $KeyTable = @(
        [PSCustomObject]@{ OfficeName="Microsoft Office 2024" ; OfficeVersion="ProPlus2024" ; Key="XJ2XN-FW8RK-P4HMP-DKDBV-GCVGB" },
        [PSCustomObject]@{ OfficeName="Microsoft Office 2021" ; OfficeVersion="ProPlus2021" ; Key="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH" },
        [PSCustomObject]@{ OfficeName="Microsoft Office 2019" ; OfficeVersion="ProPlus2019" ; Key="NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP" },
        [PSCustomObject]@{ OfficeName="Microsoft Office 2016" ; OfficeVersion="ProPlus"     ; Key="XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99" }
    )

    $KeyTable | % { if ($OfficeVersion -eq $_.OfficeVersion) { $ActivationKey = $_.Key ; $OfficeName = $_.OfficeName } }
    Write-Host " o  $OfficeName`n"


    ## Step 3: Validate KMS server connectivity
    if ($KMSserver) {
        # Optional: User input KMS activation server
        Write-Host "[+] Testing connection to input KMS server for activation..." -ForegroundColor Yellow
        $ServerList = @("$KMSserver")  
    }
    else {
        # Known public KMS activation servers
        Write-Host "[+] Determining which public KMS server to use for activation..." -ForegroundColor Yellow 
        $ServerList += @(
            'e8.us.to',
            'e9.us.to',
            'kms8.msguides.com',
            'kms9.msguides.com'
        )
    }

    # Test KMS server connection(s)
    foreach ($Server in $ServerList) {
        if (!$KMSport) { $KMsport = 1688 } # Default KMS Port
        $Connection = Test-NetConnection -ComputerName $Server -Port $KMSport
        if ($Connection.TcpTestSucceeded) { $KMS_Server = $Server ; break }
    }

    # Failed to connect to KMS server(s)
    if (!$KMS_Server) {
        cd -LiteralPath $PreActivate
        return (Write-Host "[!] Failed to connect to any KMS servers.  Try using the '-KMSserver' parameter to specify a different server.`n" -ForegroundColor Red)
    }
    Write-Host " o  ${KMS_Server}`n"


    ## Step 4: Change default license to KMS volume license (aka able to activate with KMS keys)
    Write-Host "[+] Converting retail license(s) to volume license(s)..." -ForegroundColor Yellow
    $Licenses = (Get-ChildItem "..\root\Licenses16\${OfficeVersion}VL_KMS*.xrm-ms").FullName 2>$NULL

    if ($Licenses -eq $NULL) { Write-Host " o  Skipping" }
    else { $Licenses | % { cscript /nologo ospp.vbs /inslic:$_ } }


    ## Step 5: Activate Office with KMS key
    Write-Host "`n[+] Activating Microsoft Office using KMS client key..." -ForegroundColor Yellow

    Write-Host " o  Setting KMS Hostname : " -NoNewline ; Write-Host "'${KMS_Server}'" -ForegroundColor Green
    cscript /nologo ospp.vbs /sethst:${KMS_Server}   ; Start-Sleep -Seconds 3

    Write-Host "`n o  Setting KMS Port     : " -NoNewline ; Write-Host "'${KMSport}'" -ForegroundColor Green
    cscript /nologo ospp.vbs /setprt:$KMSport        ; Start-Sleep -Seconds 3

    if (!$Office2024) {
        Write-Host "`n o  Uninstalling current product key..."
        cscript /nologo ospp.vbs /unpkey:$CurrentKey ; Start-Sleep -Seconds 3
    }

    Write-Host "`n o  Installing KMS product key..."
    cscript /nologo ospp.vbs /inpkey:$ActivationKey  ; Start-Sleep -Seconds 3

    Write-Host "`n o  Activating new product key..."
    cscript /nologo ospp.vbs /act                    ; Start-Sleep -Seconds 3


    ### Version 2.0.0 ###

    if (!$DontRollback) {

        ## Step 6: Rollback Office version to a bannerless version
        Write-Host "`n[+] Rolling back Office version to remove 'GET GENUINE OFFICE' banner..." -ForegroundColor Yellow

        $ClickToRun = “${env:ProgramFiles}\Common Files\microsoft shared\ClickToRun”
        $Continue   = $TRUE

        # Validate directory exists
        if (!(Test-Path -LiteralPath $ClickToRun)) {
            Write-Host "[!] Unable to find the '${ClickToRun}' directory. Aborting."
            $Continue = $FALSE
        }

        if ($Continue) {
            
            # Rollback method doesn't work on Office 2021 or Office 2024
            if (($OfficeVersion -eq 'ProPlus2021') -or ($OfficeVersion -eq 'ProPlus2024')) {
                Write-Host " o  Version rollback only works on Office 2016 - 2019."
            }

            # Use OfficeC2RClient to rollback version
            else {
                Write-Host " o  Rolling back to version '16.0.13801.20266'"
                Write-host " o  This might take a while..."

                cd $ClickToRun
                .\OfficeC2RClient.exe /update user updatetoversion=16.0.13801.20266
                
                # Wait for OfficeC2Client to execute and close
                Start-Sleep -Seconds 15
                Write-host " o  Wait for 'OfficeC2RClient' to finish executing before continuing."
            }

            ## Step 7: Disable automatic updates
            Write-Host "`n[+] Disabling Microsoft Office automatic updates..." -ForegroundColor Yellow
            $DisableOfficeUpdates = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'

            if (!(Test-Path $DisableOfficeUpdates)) { New-Item $DisableOfficeUpdates -Force | Out-Null }
            Set-ItemProperty -Path $DisableOfficeUpdates -Name 'enableautomaticupdates' -Value 0

            Write-Host " o  Location    : 'HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'"
            Write-Host " o  Registy Key : 'enableautomaticupdates' --> 0"
        }
    }
    else { Write-Host "`n[+] Skipping version rollback." -ForegroundColor Yellow }


    # Return to original directory post-activation
    cd -LiteralPath $PreActivate
    return (Write-Host "`n[+] Complete." -ForegroundColor Yellow)
}