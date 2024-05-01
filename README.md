# Activate-MicrosoftOffice
PowerShell tool to activate Microsoft Office Professional Plus 2016 - 2021 via KMS client keys.

## TL;DR

**Requirements:**
```
Run with elevated privileges (i.e., Administrator)
```

**One-Liner Syntax:**
```powershell
iex ([System.Net.WebClient]::new().DownloadString('https://raw.githubusercontent.com/tylerdotrar/Activate-MicrosoftOffice/main/Activate-MicrosoftOffice.ps1')); Activate-MicrosoftOffice
```

![Banner](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/bd053187-1c84-49f0-b756-12aed8ccaade)

## Description
Automatically detect and activate locally installed Microsoft Office instance with a Professional Plus 
KMS client key using a publically available KMS server.

- Supports both 32-bit and 64-bit versions of Office.
- Supports Microsoft Office 2016, 2019, and 2021.
- Script works with both Desktop PowerShell and PowerShell Core.
- Contains moderate error correction.
- **[Version 1.1.0]** Supports custom KMS servers and ports.
- **[Version 2.0.0]** Remove 'GET GENUINE OFFICE' banner via version rollback & disabled auto-updates.
- **[Version 2.0.1]** Minor visual formatting adjustments.


## Example Output
```
v2.0.0: Get-Help
```
![Get-Help](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/843f3bd5-a27c-4199-adb1-0c754d7c9ee3)


```
v2.0.1: Office version rollback.
```
![Version 2.0.1](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/735e37c8-b4e3-46e2-9b9c-41491d012487)


```
Error Correction: Not in elevated terminal.
```
![Not Elevated](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/6806500f-f6f2-4fb3-a066-91c289dd2681)


```
Error Correction: Office not installed.
```
![Failed to Find](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/405fe318-4dad-4784-ac1e-1c5d8fd3798f)


```
Error Correction: Could not connect to KMS servers.
```
![Failed to Resolve](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/6597fa4f-139c-4ebf-afb7-0aa04f5bcab4)


## Official Microsoft Office Downloads
Microsoft Office Professional Plus 2021:
- [ProPlus2021Retail.img](https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img)

Microsoft Office Professional Plus 2019:
- [ProPlus2019Retail.img](https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img)

Microsoft Office Professional Plus 2016:
- <N/A>
