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

![Output](https://cdn.discordapp.com/attachments/855920119292362802/1156658422733877298/image.png?ex=6515c599&is=65147419&hm=87a236c0363fe4644c2ac2dafeba4cce965b16d5d9fa7e36bdc9bd17e2be634b&)


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
![Get-Help](https://cdn.discordapp.com/attachments/855920119292362802/1156659521855426650/image.png?ex=6515c69f&is=6514751f&hm=de1eaf41a53770066d48d128d8a6258df3f58411dbe06259c0ae5df0f81a4c13&)

```
v2.0.1: Office version rollback.
```
![Version 2.0.1](https://cdn.discordapp.com/attachments/855920119292362802/1156672784924160180/image.png?ex=6515d2f9&is=65148179&hm=74b0d7b5da1bb23259b03359fe86ffa9a099c2365fa7f289b3f3424c557a2d7a&)


```
Error Correction: Not in elevated terminal.
```
![Not Elevated](https://cdn.discordapp.com/attachments/855920119292362802/1086409673047019550/image.png)

```
Error Correction: Office not installed.
```
![Failed to Find](https://cdn.discordapp.com/attachments/855920119292362802/1156660054322327552/image.png?ex=6515c71e&is=6514759e&hm=f96c2c6b02b10a157d267a5060f316121799659fd3c4434108e672351a16aa99&)

```
Error Correction: Could not connect to KMS servers.
```
![Failed to Resolve](https://cdn.discordapp.com/attachments/855920119292362802/1156662636566544414/image.png?ex=6515c986&is=65147806&hm=c621eb6adea81b897964a83b9fd651d88a793f4c41e56534846863edb370fadf&)


## Official Microsoft Office Downloads
Microsoft Office Professional Plus 2021:
- [ProPlus2021Retail.img](https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img)

Microsoft Office Professional Plus 2019:
- [ProPlus2019Retail.img](https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img)

Microsoft Office Professional Plus 2016:
- <N/A>
