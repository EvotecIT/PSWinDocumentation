# PSWinDocumentation - PowerShell Module

Overview of this module is available at: https://evotec.xyz/hub/scripts/pswindocumentation-powershell-module/

## Goals

Ultimate goal of this project is to have proper documentation of following services:

- Active Directory
- Microsoft Exchange
- Office 365
- Windows Server
- Windows Workstation

I'm heavily open for feature requests and people willing to create data sets. By data sets.

## Updates
- 0.0.5 / 2018.07.30
    -  fix for DefaultPassWordPoLicy MinPasswordLength (was reporting wrong value)
- 0.0.4 / 2018.07.30
    -  added domain controllers section
    -  added few verbose messages with -Verbose switch for easier debugging
    -  commented out some unused code for now (to speed up work)
- 0.0.3 / 2018.07.29
    - first "good" release

## Progress on Documentation

### Active Directory

Following is incomplete list of things that are done or are planned in near future.

- [x] Forest Summary
- [x] Forest FSMO Roles
- [x] Forest Optional Features (Recycle Bin, PAM)
- [x] Forest UPN List
- [x] Forest SPN List
- [x] Domain Summary
- [ ] Domain Controllers
    - [X] Basic information
    - [ ] Basic hardware information
- [x] Domain FSMO Roles
- [x] Domain Password Policies
- [x] Domain Group Policies
- [ ] Domain Organizational Units
    - [ ] Requires work. Currently a bit useless
- [x] Domain Priviliged Members (Groups)
- [x] Domain Administrators (All users)
- [x] Domain User Count
    - [X] Users Count Incl. System
    - [X] Users Count
    - [X] Users Expired
    - [X] Users Expired Incl. Disabled
    - [X] Users Never Expiring
    - [X] Users Never Expiring Incl. Disabled
    - [X] Users System Accounts
- [ ] Domain User List (deciding if needed)
    - [ ] Users Count Incl. System
    - [ ] Users Count
    - [ ] Users Expired
    - [ ] Users Expired Incl. Disabled
    - [ ] Users Never Expiring
    - [ ] Users Never Expiring Incl. Disabled
    - [ ] Users System Accounts

More suggestions are welcome.

### Microsoft Exchange
- [ ] Currently not defined. Feel free to make feature requests

### Microsoft Office 365
- [ ] Currently not defined. Feel free to make feature requests

### Windows Server doc
- [ ] Currently not defined. Feel free to make feature requests

### Windows Workstation doc
- [ ] Currently not defined. Feel free to make feature requests