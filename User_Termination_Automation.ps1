$LogoE = @"                                                                       
██    ██ ███████ ███████ ██████                                                         
██    ██ ██      ██      ██   ██                                                        
██    ██ ███████ █████   ██████                                                         
██    ██      ██ ██      ██   ██                                                        
 ██████  ███████ ███████ ██   ██                                                        


████████ ███████ ██████  ███    ███ ██ ███    ██  █████  ████████ ██  ██████  ███    ██ 
   ██    ██      ██   ██ ████  ████ ██ ████   ██ ██   ██    ██    ██ ██    ██ ████   ██ 
   ██    █████   ██████  ██ ████ ██ ██ ██ ██  ██ ███████    ██    ██ ██    ██ ██ ██  ██ 
   ██    ██      ██   ██ ██  ██  ██ ██ ██  ██ ██ ██   ██    ██    ██ ██    ██ ██  ██ ██ 
   ██    ███████ ██   ██ ██      ██ ██ ██   ████ ██   ██    ██    ██  ██████  ██   ████ 


 █████  ██    ██ ████████  ██████  ███    ███  █████  ████████ ██  ██████  ███    ██    
██   ██ ██    ██    ██    ██    ██ ████  ████ ██   ██    ██    ██ ██    ██ ████   ██    
███████ ██    ██    ██    ██    ██ ██ ████ ██ ███████    ██    ██ ██    ██ ██ ██  ██    
██   ██ ██    ██    ██    ██    ██ ██  ██  ██ ██   ██    ██    ██ ██    ██ ██  ██ ██    
██   ██  ██████     ██     ██████  ██      ██ ██   ██    ██    ██  ██████  ██   ████    


By Escaflowne
"@
Write-Host $LogoE -BackgroundColor Black -ForegroundColor cyan

#Import ActiveDirectory Module
Import-Module ActiveDirectory

$ADAccount = Read-Host("Type in the AD account (format Like: JSmith)")
$Email = Read-Host("Type in the email account (format Like: John.Smith@YOURDOMAIN.com)")
$Ticket = Read-Host("Enter the ticket number for this request or N/A if doesn't apply")


#Creating a Logfile. You can change the path as you please
if (!(Test-Path C:\Scripts\Logs\Terminations -PathType Container)) {mkdir C:\Scripts\Logs\Terminations}
$logFile = ("C:\Scripts\Logs\Terminations\"+ $ADAccount + " " + "TERMINATION" + " - " + (Get-Date -format MMddyyyy.HHmmss) + ".txt")
Write-Host "`r`nCreating log file: $logfile`r`n" -ForegroundColor green -BackgroundColor black
New-Item $logFile -ItemType File | Out-Null
Add-Content $logFile "User Termination as per ticket $ticket"

 try ##Disable the AD Account##
{
    Disable-ADAccount -Identity $ADAccount
	Add-Content $logFile "$ADAccount Active directory account has been disabled"
}
catch
{
    Add-Content $logFile "`r`n-----`r`nTry to disable User account failed:`r`n$($_.Exception.Message)`r`n$($_.ErrorDetails.Message)`r`n-----"
    Write-Host("There was an error disabling the user account, please check the log $logFile") -ForegroundColor White -BackgroundColor Red
}
   
try ##Force AD-Sync##
{

    Write-Host "Waiting for user to get synchronized in AzureAD" -ForegroundColor green -BackgroundColor black
    Start-Sleep -Seconds 300 #this will allow time for the synchronization process to create the user in O365
    Write-Host "`r`nSyncronizing ActiveDirectory with AzureAD" -ForegroundColor green -BackgroundColor black
    Add-Content $logFile "`r`n-----`r`nADSync"
    Start-ADSyncSyncCycle -PolicyType Delta | Out-File $logFile -Append
}
catch ##AD was not able to sync##
{
    Add-Content $logFile "`r`n-----`r`nUser was not sync with office 365:`r`n$($_.Exception.Message)`r`n-----"
    Write-Host("There was an error while trying to sync with Office 365, please check the log $logFile") -ForegroundColor White -BackgroundColor Red
}
	

try ##Office 365 process##
{ 
   Connect-AzureAD

   $AzureGroups = Get-AzureADMSGroup -All:$true
   $user = (Get-AzureADUser -ObjectId $Email)
	
	try ##Sign Off devices from AzureAD/Office 365##
	{ 
        $ObjectID = Get-AzureADUser -ObjectID $Email
        Revoke-AzureADUserAllRefreshToken -ObjectID $ObjectID.ObjectId
        Add-Content $logFile "---$ADAccount has been signed off from all devices---"
	}
    catch
	{

	    Add-Content $logFile "`r`n-----`r`nFailed to signed off user from devices:`r`n$($_.Exception.Message)`r`n$($_.ErrorDetails.Message)`r`n-----"
        Write-Host("There was an error trying to sign off $ADAccount for open sessions, please check the log $logFile") -ForegroundColor White -BackgroundColor Red
	}
    
	try ##Convert Mailbox into shared mailbox##
	{
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline
        Set-Mailbox $Email -Type Shared
	    Add-Content $logFile "---$Email has been converted to a shared mailbox---"
    }
	catch ##convertion failed##
	{
	    Add-Content $logFile "`r`n-----`r`nMailbox convertion failed:`r`n$($_.Exception.Message)`r`n$($_.ErrorDetails.Message)`r`n-----"
        Write-Host("There was an error converting the Mailbox to shared mailbox, please check the log $logFile") -ForegroundColor White -BackgroundColor Red
	}
	
	try ##Add and autoreply##
	{
		$cSAM = Read-Host "`r`nDo you want to set up an Automatic Reply? (Y/N)"
        while ($cSAM -notmatch "(y|Y|n|N)")
            {
                $cSAM = Read-Host("Please enter Y/N")
            }
        if (($cSAM -eq "Y") -or ($cSAM -eq "y"))
            {
         $InternMsg = Read-Host("Type/paste your internal message")
         $ExternMsg = Read-Host("Type/paste your external message")

        Set-MailboxAutoReplyConfiguration -Identity $Email -AutoReplyState Enabled -InternalMessage $InternMsg -ExternalMessage $ExternMsg
		Add-Content $logFile "---The Automatic Reply has been set successfully---"
            }
	}
	catch
	{
		Add-Content $logFile "`r`n-----`r`nFailed to set up the Automatic Reply:`r`n$($_.Exception.Message)`r`n$($_.ErrorDetails.Message)`r`n-----"
        Write-Host("There was an error while trying to set up the Automatic Reply, please check the log $logFile") -ForegroundColor White -BackgroundColor Red
	}
	
	try ##Give access to other users to this shared mailbox##
	{
		$NewOwner = Read-Host "`r`nDo you need to provide someone access to this shared mailbox? (Y/N)"
        while ($NewOwner -notmatch "(y|Y|n|N)")
            {
                $NewOwner = Read-Host("Please enter Y/N")
            }
         if (($NewOwner -eq "Y") -or ($NewOwner -eq "y"))
            {
              $AssignPermission = Read-Host("Type the email of the user that will have access to this shared mailbox")

              Add-MailboxPermission -Identity $Email -User $AssignPermission -AccessRights FullAccess -InheritanceType All
			  Add-Content $logFile "---$AssignPermission has now access to the shared mailbox $Email---"
            }
	}
	catch
	{
		Add-Content $logFile "`r`n-----`r`nFailed to provide access to the Shared Mailbox:`r`n$($_.Exception.Message)`r`n$($_.ErrorDetails.Message)`r`n-----"
        Write-Host("There was an error while trying to provide access to $AssignPermission to the Shared Mailbox $email, please check the log $logFile") -ForegroundColor White -BackgroundColor Red
	}
	
	try ##Remove the user License##
	{
   
        connect-msolservice
        (get-MsolUser -UserPrincipalName $email).licenses.AccountSkuId |
        foreach{
                 Set-MsolUserLicense -UserPrincipalName $email -RemoveLicenses $_
               }
        Add-Content $logFile "---Office 365 License has been removed---"
	}
	catch ##Remove the user License failed##
	{
	    Add-Content $logFile "`r`n-----`r`nFailed to remove user license:`r`n$($_.Exception.Message)`r`n$($_.ErrorDetails.Message)`r`n-----"
        Write-Host("There was an error while removing the license from the user account, please check the log $logFile") -ForegroundColor White -BackgroundColor Red
	}
}
catch ##connection to AzureAD failed##
{
        Add-Content $logFile "`r`n-----`r`nConnection to AzureAD Failed:`r`n$($_.Exception.Message)`r`n-----"
        Write-Host("There was an error connecting to AzureAD / Office365, please check the log $logFile") -ForegroundColor White -BackgroundColor Red
		
}
 finally
{
    Write-Host "The process has finished. Please check $logfile for details. Gracias =)"
    Get-PSSession | Remove-PSSession -ErrorAction SilentlyContinue
    invoke-Item $logFile
    [void](Read-Host 'Press Enter to exit…')
}

