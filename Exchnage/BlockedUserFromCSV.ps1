############################################################################
#
#
#  
#  Block user mailboxes from csv file
#
#  Author: 		Viktor Kucher
#  Email:  		viktor.kucher@gmail.ua
#  Date created:   	02/04/2014
#  Version: 		1.0 
#  Description: Set mailbox parameters that make mailbox is not accesible to send or recieve mail.
#		List all emails to block put in csv file like this:
#
#
#	User
#	staff1@company.ua
#	staff2@company.ua
#	staff3@company.ua
# 
#		It usefull when staff leave company. In order to reach it script set certain paramters. 
#
#   	1. Remove from distribution group
#	2. Hidden from GAL.                    
#  	3. Disable address policy.             
#   	4. Change smtp addresses.              
#   	5. Max Send Size 0Mb.                  
#   	6  Change email addresses.             
#	7  Accept messages only from postmaster
#                                                


#####   Change parameters if you need
#
$LogFile 	= 'c:\scripts\Logs\BlockUsers.log'
$postmaster	= 'postmaster@company.ua'
$UsersList	= 'c:\scripts\BlockMailboxes\BlockUsersList.csv'
#
#####



$Users = import-csv -path $UsersList   -delimiter ';' 

foreach ($UserData in $Users)
{
$User = $UserData.User

     $LogMessage = $(Get-Date -Format "dd.MM.yyyy HH:mm:ss") +' Starting processing: '+ $User + "`r`n"
     Write-Host $LogMessage
	try
	{
     
	 $UserMailBox = Get-Mailbox -identity $User -ErrorAction Stop
	 
	 
##############  Remove from distribution group

	write-host 'Start to remove from distribution mail group'
	ForEach ($group in Get-DistributionGroup)
	{
	   ForEach ($Member in Get-DistributionGroupMember -ResultSize Unlimited -identity $Group | Where { $_.Alias –eq $UserMailBox.Alias })
	   {
		write-host $group.alias $group.count
		Remove-DistributionGroupMember -Identity $group.alias -Member $Member.Alias -Confirm:$False -BypassSecurityGroupManagerCheck -ErrorAction Stop
	        $LogMessage += $(Get-Date -Format "dd.MM.yyyy HH:mm:ss     ") +' remove '+ $user +' from ' + $group.name + "`r`n"

	   }
	}



###############  Set parameters for block 


        $NewSmtpList = @()

	foreach ($address in $UserMailBox.EmailAddresses)
	{
	$prefixMailRandom = Get-Random
	$NewSmtp = $address.smtpaddress.substring(0,$address.smtpaddress.LastIndexOf('@')) +  $prefixMailRandom + $address.smtpaddress.substring($address.smtpaddress.LastindexOf('@'), $address.smtpaddress.length - $address.smtpaddress.LastindexOf('@') ) 
	
	write-host ' Changing smtp address: ' $address.smtpaddress ' --> ' $NewSmtp

	$NewSmtpList += $NewSmtp 
	$LogMessage += $(Get-Date -Format "dd.MM.yyyy HH:mm:ss     ") +' Changing smtp address: '+  $address.smtpaddress +' --> '+  $NewSmtp +"`r`n"
	}

	Set-Mailbox -Identity $User -HiddenFromAddressListsEnabled $true  -EmailAddressPolicyEnabled $false -MaxSendSize 0Mb  -EmailAddresses $NewSmtpList  -AcceptMessagesOnlyFrom $postmaster  -ErrorAction Stop

###############  Save to log

	}

	catch
	{
            write-host "Caught an exception:" -ForegroundColor Red
	    write-host "Exception Type: $($_.Exception.GetType().FullName)" -ForegroundColor Red
	    write-host "Exception Message: $($_.Exception.Message)" -ForegroundColor Red
            $LogMessage += $(Get-Date -Format "dd.MM.yyyy HH:mm:ss     ") +$($_.Exception.Message)+ "`r`n"
	    Add-Content -Path $LogFile $LogMessage
	   # break	
	}

	$LogMessage += $(Get-Date -Format "dd.MM.yyyy HH:mm:ss") +' Finish processing: '+ $user + "`r`n"
      	Add-Content -Path $LogFile $LogMessage
}
      	