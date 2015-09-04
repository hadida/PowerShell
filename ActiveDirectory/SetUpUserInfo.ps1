############################################################################
#
#  
#  Set Custom Attributes
#  Author: Viktor Kucher
#  Date created:   22/01/2014
#  Version: 1.0 
#  Description: 
#






#Add Exchange 2010 snapin if not already loaded
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.E2010"}))
{
	Write-Verbose "Loading Exchange 2010 Snapin"
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.E2010 -ErrorAction SilentlyContinue
}


###############  Change path #####################


$InFile  	= "D:\vkucher\MailboxUsersNew.csv"
$InFileUtf8  	= "D:\vkucher\MailboxUsersNewUtf8.csv"
$LogUsersInfo	= "D:\vkucher\LogUsersInfo.log"

##################################################



Get-Content $InFile | Set-Content -Encoding UTF8 $InFileUtf8
$Users = import-csv $InFileUtf8 -delimiter ';'

$i = 1

foreach ($user in $Users)

{


$FirstNameNew 	= ($user.FirstNameNew).TrimEnd()
$LastNameNew  	= ($user.LastNameNew).TrimEnd()
$DisplayNameNew	= $LastNameNew +' '+$FirstNameNew
$Account	= $user.SamAccountName


$CurrentMailbox = get-user -Identity $Account 






	try
	{
	Write-Host  'User: '$i $Account  $FirstnameNew  $LastnameNew  $DisplayNameNew
        Set-user -Identity $Account  -FirstName $FirstNameNew -LastName $LastNameNew -DisplayName $DisplayNameNew  -ErrorAction Stop 
	$LogMessage = $CurrentMailBox.SamAccountName +';'+$CurrentMailBox.FirstName+';'+$CurrentMailBox.LastName+';'+$CurrentMailBox.DisplayName+';'+$FirstNameNew+';'+$LastNameNew+';'+$DisplayNameNew
	Add-content -Path $LogUsersInfo $LogMessage
	}
	catch
	{
	$ErrorMessage = $_.Exception.Message
	$FailedItem = $_.Exception.ItemName
	write-host $Account  'Error to set new parameters' $ErrorMessage $FailedItem
	$LogMessage = $Account + ';Error to set new parameters'+';'+$ErrorMessage+';'+ $FailedItem
	Add-content -Path $LogUsersInfo $LogMessage
	}

	$i += 1




}


