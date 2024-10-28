function Activate-MicrosoftOffice {
#.SYNOPSIS
# Activate Microsoft Office 2016 - 2021 via free Professional Plus KMS client keys 
# Arbitrary Version Number:  2.0.2
# Author:  Tyler McCann (@tylerdotrar)
#
#.DESCRIPTION
# Automatically detect and activate a locally installed Microsoft Office instance with a Professional Plus KMS client key,
# using either a publicly available KMS server (default) or a specfied KMS server.  This script supports both 32-bit and
# 64-bit versions of Microsoft Office 2016, 2019, and 2021.
#
# As of version 2.0.0, this script will rollback the Microsoft Office version to '16.0.13801.20266' and set a registry key
# to disable automatic updates.  This should prevent the annoying 'GET GENUINE OFFICE' banner that plagued earlier versions
# of this script when using these KMS client keys.
# 
# Usage Notes:
# o This script must run with elevated privileges (i.e., as Administrator)
# o This script contains a list of known public KMS servers and will automatically
#   attempt to use them if the user doesn't specify a specific KMS server and/or port.
# o If you don't want to disable automatic updates or rollback your Office version, 
#   run this script with the '-DontRollback' parameter.
#
# Parameters:
#   -KMSserver     -->  Target domain/IP of a specified KMS server
#   -KMSport       -->  Target port of a specified KMS server
#   -DontRollback  -->  Do not rollback Office version and leave automatic updates enabled
#   -Help          -->  Return Get-Help information 
#
# Official Microsoft Office Downloads:
#   [+] Microsoft Office 2021 (Professional Plus):
#    o  Download Link: https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img
#
#   [+] Microsoft Office 2019 (Professional Plus):
#    o  Download Link: https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img
#
#   [+] Microsoft Office 2016 (Professional Plus):
#    o  Download Link: https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-us/ProPlusRetail.img
#
#.LINK
# https://github.com/tylerdotrar/Activate-MicrosoftOffice

    
    Param ( 
        [string]$KMSserver,
        [int]   $KMSport,
        [switch]$DontRollback,
        [switch]$Help
    )


    # Return Get-Help Information
    if ($Help) { return (Get-Help Activate-MicrosoftOffice) }


    # Determine if user has elevated privileges
    $User    = [Security.Principal.WindowsIdentity]::GetCurrent();
    $isAdmin = (New-Object Security.Principal.WindowsPrincipal $User).IsInRole([Security.Principal.WindowsBuiltinRole]::Administrator)
    if (!$isAdmin) { return (Write-Host '[-] This function requires elevated privileges.' -ForegroundColor Red) }
    

    ## Step 1: Verify Office is installed
    Write-Host '[+] Determining if Office 64-bit or 32-bit version is installed...' -ForegroundColor Yellow
    $PreActivate  = "$PWD"
    $OfficeDir    = "$env:ProgramFiles\Microsoft Office\Office16"
    $OfficeDirx86 = "${env:ProgramFiles(x86)}\Microsoft Office\Office16"


    if (Test-Path -LiteralPath $OfficeDir)        { cd -LiteralPath $OfficeDir    ; Write-Host " o  64-bit`n" }
    elseif (Test-Path -LiteralPath $OfficeDirx86) { cd -LiteralPath $OfficeDirx86 ; Write-host " o  32-bit`n" }
    else { return (Write-Host "[-] Failed to find Microsoft Office.`n" -ForegroundColor Red) }

    
    ## Step 2: Determine version of Office installed
    Write-Host "[+] Determining version of Office to activate..." -ForegroundColor Yellow
    $OfficeStatus  = cscript /nologo ospp.vbs /dstatus
    $OfficeVersion = ($OfficeStatus | Select-String -Pattern "ProPlus(\d{4})?").Matches.Value
    $CurrentKey    = (($OfficeStatus | Select-String -Pattern "product key:") -as [string]).Split(' ')[-1]


    if (!$OfficeVersion) {
        cd -LiteralPath $PreActivate
        return (Write-Host "[-] Initial Office license not detected.  Please first-time open a Microsoft Office executable or reboot your computer.`n" -ForegroundColor Red)
    }
    

    # Microsoft Office KMS Client Key Table
    $KeyTable = @(
        [PSCustomObject]@{OfficeName="Microsoft Office 2021"; OfficeVersion="ProPlus2021"; Key="FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH"},
        [PSCustomObject]@{OfficeName="Microsoft Office 2019"; OfficeVersion="ProPlus2019"; Key="NMMKJ-6RK4F-KMJVX-8D9MJ-6MWKP"},
        [PSCustomObject]@{OfficeName="Microsoft Office 2016"; OfficeVersion="ProPlus"    ; Key="XQNVK-8JYDB-WJ9W3-YJ8YR-WFG99"}
    )


    $KeyTable | % { if ($OfficeVersion -eq $_.OfficeVersion) { $ActivationKey = $_.Key ; $OfficeName = $_.OfficeName } }
    Write-Host " o  $OfficeName`n"


    ## Step 3: Validate KMS server connectivity
    if ($KMSserver) { Write-Host "[+] Testing connection to input KMS server for activation..." -ForegroundColor Yellow     }
    else            { Write-Host "[+] Determining which public KMS server to use for activation..." -ForegroundColor Yellow }


    # Optional: User input KMS activation server
    if ($KMSserver) { $ServerList = @("$KMSserver") }


    # Known public KMS activation servers
    else {
        $ServerList += @(
            'e8.us.to',
            'e9.us.to',
            'kms8.msguides.com',
            'kms9.msguides.com'
        )
    }


    # Test KMS server connection(s)
    foreach ($Server in $ServerList) {
        # Default KMS Port: 1688
        if (!$KMSport) { $KMsport = 1688 }
        $Connection = Test-NetConnection -ComputerName $Server -Port $KMSport

        if ($Connection.TcpTestSucceeded) { $KMS_Server = $Server ; break }
    }


    # Failed to connect to KMS server(s)
    if (!$KMS_Server) {
        cd -LiteralPath $PreActivate
        return (Write-Host "[-] Failed to connect to any KMS servers.  Try using the '-KMSserver' parameter to specify a different server.`n" -ForegroundColor Red)
    }
    Write-Host " o  ${KMS_Server}`n"


    ## Step 4: Change default license to KMS volume license (aka able to activate with KMS keys)
    Write-Host "[+] Converting retail license(s) to volume license(s)..." -ForegroundColor Yellow
    $Licenses = (Get-ChildItem "..\root\Licenses16\${OfficeVersion}VL_KMS*.xrm-ms").FullName 2>$NULL


    if ($Licenses -eq $NULL) { Write-Host " o  Skipping" }
    else { $Licenses | % { cscript /nologo ospp.vbs /inslic:$_ } }


    ## Step 5: Activate Office with KMS key
    Write-Host "`n[+] Activating Microsoft Office using KMS client key..." -ForegroundColor Yellow

    Write-Host "`n o  Setting KMS Hostname : '${KMS_Server}'"
    cscript /nologo ospp.vbs /sethst:${KMS_Server}  ; Start-Sleep -Seconds 3

    Write-Host "`n o  Setting KMS Port     : $KMSport"
    cscript /nologo ospp.vbs /setprt:$KMSport       ; Start-Sleep -Seconds 3

    Write-Host "`n o  Uninstalling current product key..."
    cscript /nologo ospp.vbs /unpkey:$CurrentKey    ; Start-Sleep -Seconds 3

    Write-Host "`n o  Installing KMS product key..."
    cscript /nologo ospp.vbs /inpkey:$ActivationKey ; Start-Sleep -Seconds 3

    Write-Host "`n o  Activating new product key..."
    cscript /nologo ospp.vbs /act                   ; Start-Sleep -Seconds 3


    ### Version 2.0.0 ###
    if (!$DontRollback) {

        # Step 6: Rollback Office version to a bannerless version
        Write-Host "`n[+] Rolling back Office version to remove 'GET GENUINE OFFICE' banner..." -ForegroundColor Yellow

        $ClickToRun = “$env:ProgramFiles\Common Files\microsoft shared\ClickToRun”
        $Continue   = $TRUE

        # Validate directory exists
        if (!(Test-Path -LiteralPath $ClickToRun)) {
            Write-Host "[-] Unable to find the '$ClickToRun' directory. Aborting."
            $Continue = $FALSE
        }

        if ($Continue) {
            
            # Rollback method doesn't work on Office 2021
            if ($OfficeVersion -eq 'ProPlus2021') {
                Write-Host " o  Version rollback only works on Office 2016 - 2019.`n"
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

                # Removed first-time launch false negatives
                <#
                while (Get-Process -Name OfficeC2RClient 2>$NULL) { $SuccessfullyExecuted = $TRUE }
                if (!$SuccessfullyExecuted) {  Write-Host "[-]  Failed to launch 'OfficeC2RClient' executable." }
                #>
            }
            

            # Step 7: Disable automatic updates
            Write-Host "`n[+] Disabling Microsoft Office automatic updates..." -ForegroundColor Yellow
            $DisableOfficeUpdates = 'Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'

            if (!(Test-Path $DisableOfficeUpdates)) { New-Item $DisableOfficeUpdates -Force | Out-Null }
            Set-ItemProperty -Path $DisableOfficeUpdates -Name 'enableautomaticupdates' -Value 0

            Write-Host "  o  Location    : 'HKLM\SOFTWARE\Policies\Microsoft\office\16.0\common\officeupdate'"
            Write-Host "  o  Registy Key : 'enableautomaticupdates' --> 0"
        }
    }
    else { Write-Host "`n[+] Skipping version rollback." -ForegroundColor Yellow }


    # Return to original directory post-activation
    cd -LiteralPath $PreActivate
    return (Write-Host "`n[+] Complete." -ForegroundColor Yellow)
}
