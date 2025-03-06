# Activate-MicrosoftOffice
**PowerShell tool to activate **Microsoft Office 2016 - 2021** via static Professional Plus KMS client keys.**

_(Note: this does not work on anything beyond Office 2021 due to Microsoft's transition to cloud-based Office 365)_


## TL;DR

![image](https://github.com/user-attachments/assets/32d3482b-c9ee-4e75-8c34-c1a0a3f3d720)

### Requirements:

> 1.  Run script with elevated privileges (i.e., as Administrator).
> 2.  Internet connectivity for online key activation (only needed for the initial script execution).

### Usage:

```powershell
# One-Liner: Execute Script w/ Default Parameters
irm https://tylerdotrar.github.io/Activate-MicrosoftOffice | iex; Activate-MicrosoftOffice
```
```powershell
# Step-by-Step: Load Script into the Current Session
$ScriptContents = Invoke-RestMethod -Uri 'https://tylerdotrar.github.io/Activate-MicrosoftOffice'
Invoke-Expression -Command $ScriptContents

# Activate Office w/ Default Parameters
Activate-MicrosoftOffice

# Activate Office, but do NOT disable automatic updates (might result in 'GET GENUINE OFFICE' banner)
Activate-MicrosoftOffice -DontRollBack
```


## Official Microsoft Office Downloads

While not a hard requirement, it is recommended to use one of the **official** Microsoft Office download links listed below:

- Microsoft Office 2021 (Professional Plus)  →  [ProPlus2021Retail.img](https://officecdn.microsoft.com/db/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2021Retail.img)
- Microsoft Office 2019 (Professional Plus)  →  [ProPlus2019Retail.img](https://officecdn.microsoft.com/pr/492350f6-3a01-4f97-b9c0-c7c6ddf67d60/media/en-us/ProPlus2019Retail.img)
- Microsoft Office 2016 (Professional Plus)  →  [ProPlusRetail.img](https://officecdn.microsoft.com/db/492350F6-3A01-4F97-B9C0-C7C6DDF67D60/media/en-us/ProPlusRetail.img)


## Description

_"What are KMS keys?"_
> Essentially, Key Management System (KMS) client keys are static enterprise keys designed to handle large-scale activation within an organization.  They generally rely on a local KMS activation server, allowing organizations to activate multiple clients on their network via volume licensing, rather than each individual client connecting to Microsoft for activiation.
>
> In contrast, consumer keys (such as retail or OEM) are essentially dynamic, one-time-use keys intended for small-scale, individual user activation.  These are what you usually use when you activate your product.

This project modifies local versions of Microsoft Office to utilize Microsoft's static KMS client keys rather than one-time-use consumer keys.  The Office instance is then activated with a Professional Plus KMS client key via a publicly available KMS activation server.

This PowerShell script only supports Microsoft Office 2016, 2019, and 2021 -- and has been tested on Windows 10 and Windows 11.

- **Script Features:**
  - Supports both 32-bit and 64-bit versions of Office.
  - Supports Microsoft Office 2016, 2019, and 2021.
  - Script works with both Desktop PowerShell and PowerShell Core.
  - Contains moderate error correction.
  - **[Version 1.1.0]** Supports custom KMS servers and ports.
  - **[Version 2.0.0]** Remove `GET GENUINE OFFICE` banner via version rollback & disabled auto-updates.
  - **[Version 2.0.1]** Minor visual formatting adjustments.


## Example Output(s)
```
Get-Help information (as of v2.0.2).
```
![Get-Help](https://github.com/user-attachments/assets/1f9cdd20-b50a-4d57-b5cc-251bc842b428)


```
Office version rollback introduced in v2.0.0.
```
![Version 2.0.1](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/735e37c8-b4e3-46e2-9b9c-41491d012487)

```
Error Correction:  Not executed in elevated terminal.
```
![Not Elevated](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/6806500f-f6f2-4fb3-a066-91c289dd2681)

```
Error Correction:  Microsoft Office not installed.
```
![Failed to Find](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/405fe318-4dad-4784-ac1e-1c5d8fd3798f)

```
Error Correction:  Could not connect to KMS server(s).
```
![Failed to Resolve](https://github.com/tylerdotrar/Activate-MicrosoftOffice/assets/69973771/6597fa4f-139c-4ebf-afb7-0aa04f5bcab4)
