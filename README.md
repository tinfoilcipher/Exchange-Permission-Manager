# ExchangePermissionManager
Exchange User Permission Manager

## OUTLINE
The purpose of this script is to provide a single console to manage Exchange Full Access and Send-As permissions for both
On Premise Exchange (2010+) and Office 365 Exchange Online without needing to use the cumbersome process of entering the
commands manually or using the slow GUI methods. Supports input of batch entries via prompt.

## PREREQUISITES
ActiveDirectory Module (For On Prem Send-As)
Exchange PowerShell Modules
Properly configured On Prem Exchange Powershell, see **USE - ON PREMISE**

## USE - GENERAL
On execution, you will be prompted to chose Exchange Online or On Premise with 1 or 2 as an input. This selects the correct
connection.<br>
<br>
After connection, you will be asked what function you want to perform from 1 -5. Entry outside of this range will cause an abort.<br>
<br>
Write-Host "Exchange Online or On Premise Exchange?"
1. Grant Full Mailbox Access
2. Grant Full Mailbox Access w/out Automapping
3. Remove Full Mailbox Access
4. Grant Send-As Rights On-Prem Only
5. Grant Send-As Rights Exchange Online Only
<br>
For either choice you'll then be asked to provide a list of users to grant/remove permissions FOR. Input can be entered in the
following formats<br>

|no_headings_needed   | 
|---------------------| 
| user1,user2,user3   | 
| user1, user2, user3 | 
| user1; user2, user3 | 

**Any combination of the above are valid**<br>

Either the Exchange Aliases, primary SMTPs or userPrincipalNames should be valid inputs.<br>
<br>
You will then be asked for a list of mailboxes to apply these changes TO, input is in the same format as above<br>
<br>
The function will then be called and the permission applied.<br>
1. Owner (Full Access)
2. Editor (Read Write - Edit All)
3. Author (Read Write - Edit Own)
4. Reviewer (Read Only)
5. Contributor (Write Only)

## USE - ON PREMISE
Should work with any Exchange Server on premise where the following is true:<br>
1. Autodiscover is configured<br>
2. Remote Powershell is enabled and you have permissions<br>
3. The server has an internally resolvable FQDN<br>
Though if the above conditions aren't met, your Exchange is probably in a seriously unhealthy state to begin with.<br>
<br>
To execute, enter the FQDN of your CAS as the $strOnPremServer Constant and execute.enter the following as constants in the
Constants block<br>

1. $strOnPremServer. Set to the FQDN of your CAS, prefixed with https:// (for use in making an Exchange PowerShell Connection)
2. $strNetBIOS. Set to the NETBIOS name of your domain, for use in writing Send-As as an Active Directory permission.

## USE - OFFICE 365
No change to constants are needed, you will however be prompted to enter the credentials for a user that has Exchange Admin
rights to the tenancy you are connecting to.
