#Exchange Permission Manager - tinfoilcipher 2018

#Constants (On Prem Only)
$strOnPremServer = "https://server.domain.tld"
$strNetBIOS = "DOMAIN"
$strURI = ($strOnPremServer + "/powershell-liveid/")

#Functions
Function Clean-and-Split($Dirty) {
	$Dirty = $Dirty.Replace(" ", "")
	$Dirty = $Dirty.Replace(";", ",")
	$Dirty = $Dirty.Split(",")
}
Function Connect-365 {
	$strCredential = Get-Credential
	$365Esession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $strCredential -Authentication Basic -AllowRedirection
	Import-PSSession $365Esession
}
Function Connect-OnPrem {
	$OnPremSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $strURI -Authentication Kerberos
	Import-PSSession $OnPremSession
	Import-Module ActiveDirectory
}
Function Grant-FullAuto {
	ForEach ($Subject In $Subjects) {
		ForEach ($User In $Users) {
			Add-MailboxPermission -Identity $Subject -User $User -AccessRights Fullaccess -InheritanceType all
		}
	}
}
Function Grant-Full {
	ForEach ($Subject In $Subjects) {
		ForEach ($User In $Users) {
			Add-MailboxPermission -Identity $Subject -User $User -AccessRights Fullaccess -InheritanceType all -AutoMapping $false
		}
	}

}
Function Remove-Full {
	ForEach ($Subject In $Subjects) {
		ForEach ($User In $Users) {
			Remove-MailboxPermission -Identity $Subject -User $User -AccessRights Fullaccess -confirm:$false
		}
	}
}
Function Grant-SendAsOnPrem {
	Foreach ($Subject in $Subjects) {
		Foreach ($User in $Users) {
			Get-Mailbox -Identity $Subject | Add-ADPermission -User ($strNetBIOS + "\" + $User) -Extendedrights "Send As"
		}
	}
}
Function Grant-SendAsExchangeOnline {
	Foreach ($Subject in $Subjects) {
		Foreach ($User in $Users) {
			Add-RecipientPermission $User -AccessRights SendAs -Trustee $Subject
		}
	}
}

#Select Connection
Clear-Host
Write-Host "Exchange Online or On Premise Exchange?"
Write-Host "1 - Exchange Online"
Write-Host "2 - Exchange OnPrem"
$strExchange = Read-Host ">"

If ($strExchange -eq "1"){
	Connect-365
}
ElseIf ($strExchange -eq "2"){
	Connect-OnPrem
}

#Select Function
Clear-Host
Write-Host "Exchange Online or On Premise Exchange?"
Write-Host "1 - Grant Full Mailbox Access"
Write-Host "2 - Grant Full Mailbox Access w/out Automapping"
Write-Host "3 - Remove Full Mailbox Access"
Write-Host "4 - Grant Send-As Rights On-Prem Only"
Write-Host "5 - Grant Send-As Rights Exchange Online Only"
$strOption = Read-Host ">"
Clear-Host

#Input user(s)
Write-Host "Enter the alias(es) of the mailbox(es) for users who require access to be granted or revoked (the users"
Write Host "you're granting access TO or revoking access FROM)."
$Users = Read-Host "Separate multiple LoginIDs with commas"
Clear-Host

#Input target(s)
Write-Host "Enter the alias(es) of the target mailboxes (the mailboxes you're applying permissions TO)"
$Subjects = Read-Host "Separate multiple LoginIDs with commas"
Clear-Host

#Build arrays
$Users = Clean-and-Split($Users)
$Subjects = Clean-and-Split($Subjects)

#Perform Action
If ($strOption -eq "1") {
	Grant-FullAuto
}
ElseIf ($strOption -eq "2") {
	Grant-Full
}
ElseIf ($strOption -eq "3") {
	Remove-Full
}
ElseIf ($strOption -eq "4") {
	Grant-SendAsOnPrem
}
ElseIf ($strOption -eq "5") {
	Grant-SendAsExchangeOnline 
}
Else {
	Clear-Host
	Write-Host "Invalid Option Selected. No Action Taken. Press Return To Exit."
	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	$host.SetShouldExit(0)
	exit
}

#Success
Clear-Host
Write-Host "Permissions Modified As Requested. Press Return To Exit."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
$host.SetShouldExit(0)
exit

#END.
