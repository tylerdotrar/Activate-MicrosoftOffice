# Activate-MicrosoftOffice
PowerShell tool to activate Microsoft Office Professional Plus 2016 - 2021 via KMS client keys.

## TL;DR

**Requirements:**
```
Run with elevated privileges (i.e., Administrator)
```

**One-Liner Syntax:**
```powershell
iex ([System.Net.WebClient]::new().DownloadString('https://raw.githubusercontent.com/tylerdotrar/Activate-MicrosoftOffice/main/Activate-MicrosoftOffice.ps1'))
```

![Output](https://cdn.discordapp.com/attachments/855920119292362802/1086394905296916561/image.png)

## Description
Automatically detect and activate locally installed Microsoft Office instance with a Professional Plus 
KMS client key using a publically available KMS server.
- Supports both 32-bit and 64-bit
- Supports Microsoft Office 2016, 2019, and 2021
- Script works in Desktop PowerShell and PowerShell Core
- Moderate error correction

![Error Correction](https://cdn.discordapp.com/attachments/855920119292362802/1086403306236162138/image.png)

## Get-Help
![Get-Help](https://cdn.discordapp.com/attachments/855920119292362802/1086407625110999060/image.png)

## Official Microsoft Office Downloads
Microsoft Office Professional Plus 2021:
- [ProPlus2021Retail.img](https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img)

Microsoft Office Professional Plus 2019:
- [ProPlus2019Retail.img](https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img)

Microsoft Office Professional Plus 2016:
- <N/A>
