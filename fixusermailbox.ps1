Install-Module -Name ExchangeOnlineManagement
Import-Module -name ExchangeOnlineManagement

Set-ExecutionPolicy RemoteSigned

install-module exchangeonlinemanagement

Connect-ExchangeOnline
Get-Mailbox useremailaddress@office.com | FL SingleItemRecoveryEnabled,RetainDeletedItemsFor
Get-CASMailbox useremailaddress@office.com | FL EwsEnabled,ActiveSyncEnabled,MAPIEnabled,OWAEnabled,ImapEnabled,PopEnabled
Get-Mailbox useremailaddress@office.com | FL LitigationHoldEnabled,InPlaceHolds
Get-OrganizationConfig | FL InPlaceHolds
Get-Mailbox useremailaddress@office.com | FL DelayHoldApplied,DelayReleaseHoldApplied
Get-MailboxFolderStatistics useremailaddress@office.com -FolderScope RecoverableItems | FL Name,FolderAndSubfolderSize,ItemsInFolderAndSubfolders
Set-Mailbox useremailaddress@office.com -SingleItemRecoveryEnabled $true
Set-Mailbox useremailaddress@office.com -LitigationHoldEnabled $true
Start-ManagedFolderAssistant -Identity "useremailaddress@office.com"
Get-Mailbox "useremailaddress@office.com" | FL DelayHoldApplied,DelayReleaseHoldApplied
Set-Mailbox "useremailaddress@office.com" -RemoveDelayHoldApplied
Set-Mailbox "useremailaddress@office.com" -RemoveDelayReleaseHoldApplied
Start-ManagedFolderAssistant -Identity "useremailaddress@office.com"
Set-Mailbox –Identity "useremailaddress@office.com" -RetainDeletedItemsFor 14

