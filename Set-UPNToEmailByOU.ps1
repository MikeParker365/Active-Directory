<#
.SYNOPSIS
Updates the UserPrincipalName attribute for all users under a specified OU. This can be done on an OU by OU basis or at a top level OU which will also update all the users in any child OUs.

Author/Copyright:    Mike Parker - All rights reserved
Email/Blog/Twitter:  mike@mikeparker365.co.uk | www.mikeparker365.co.uk | @MikeParker365

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS CODE REMAINS WITH THE USER.

.NOTES
Version 1.0, 6th January, 2016

Revision History
---------------------------------------------------------------------
1.0	Initial release
	
.DESCRIPTION
Updates the UserPrincipalName attribute for all users under a specified OU. This can be done on an OU by OU basis or at a top level OU which will also update all the users in any child OUs.
	
.PARAMETER OU
Specifies the distinguished name of the Org. Unit where you would like to update the UPNs for all users.

.PARAMETER TxtLog
Specifies the path where you would like to save a verbose text log of changes made.

.PARAMETER CsvLog
Specifies the path where you would like to save a CSV output of the users updated.



.LINK
http://www.mikeparker365.co.uk

.EXAMPLE
PS C:\> Set-UPNToEmailByOU.ps1 -OU "OU=Cloud Users,DC=Cloud,DC=local" -TxtLog "C:\Logs\"

.EXAMPLE
PS C:\> Set-UPNToEmailByOU.ps1 -OU "OU=IT Dept,OU=Cloud Users,DC=Cloud,DC=local" -TxtLog "C:\Logs\" -CsvLog "C:\Logs\"

    
#>

Param (
 
    [Parameter( Mandatory=$true )] 
    [string]$OU, 
    [Parameter( Mandatory=$true )] 
    [string]$TxtLog,
    [Parameter( Mandatory=$false )] 
    [string]$CsvLog

) 

############################################################################
# Global Variables Start 
############################################################################

$LogFolder = "C:\CloudBusiness\UPN Update Logs\";

$dateForOutput = get-date -Format ddMMyyyy
$logCSV = $CsvLog + "UPN_Update_" + $dateForOutput + ".csv";
$logFile = $txtLog + "UPNLogVerbose" + $dateForOutput + ".txt";

$recordsProcessed = 0;
$recordsUpdated = 0;
$recordsCorrect = 0;
$recordsError = 0;
$start = Get-Date

############################################################################
# Global Variables End
############################################################################

############################################################################
# Functions Start 
############################################################################


#Defines functions to be used to output progress. Use these within the script to save time.
function ShowError ($msg){Write-Host "`n";Write-Host -ForegroundColor Red $msg; LogErrorToFile $msg }
function ShowSuccess($msg){Write-Host "`n";Write-Host -ForegroundColor Green $msg; LogToFile ($msg)}
function ShowProgress($msg){Write-Host "`n";Write-Host -ForegroundColor Cyan $msg; LogToFile ($msg)}
function ShowInfo($msg){Write-Host "`n";Write-Host -ForegroundColor Yellow $msg; LogToFile ($msg)}
function LogToFile ($msg){$msg |Out-File -Append -FilePath $LogFile -ErrorAction:SilentlyContinue;}
function LogSuccessToFile ($msg){"Success: $msg" |Out-File -Append -FilePath $LogFile -ErrorAction:SilentlyContinue;}
function LogErrorToFile ($msg){"Error: $msg" |Out-File -Append -FilePath $LogFile -ErrorAction:SilentlyContinue;}


############################################################################
# Functions end 
############################################################################

############################################################################
# Required Modules Start 
############################################################################


ShowProgress "Calling the ActiveDirectory Cmdlets";

try{
    $error.Clear()
    Import-Module ActiveDirectory    
    }
catch{
    ShowError "There was an error importing the AD Cmdlets"
    ShowError "$error"
    }
finally{
    if(!$error){
        ShowSuccess "Successfully loaded AD Cmdlets"
        }
    else{
        ShowError "There was an error importing the AD Cmdlets"
        ShowError "The error logged was - $error"
        }
    }

############################################################################
# Module Import End
############################################################################ 

############################################################################
# Script start   
############################################################################

ShowProgress "Script started at $start";

$error.clear()

ShowProgress "Looking for users in OU: $OU"

$Users = Get-ADUser -Filter * -Properties * -SearchBase $OU
$UserCount = $Users.Count

$Users | Foreach { 
	$error.clear()
	$recordsProcessed++

	Write-Progress -Activity "Processing UPN Updates..." -Status "User $recordsProcessed of $userCount" -PercentComplete ($recordsProcessed / $userCount * 100)

	$displayName = $_.Name
	$OldUpn = $_.UserPrincipalName
	$NewUpn = $_.mail

	if($OLDUpn.ToLower() -eq $NewUPN.ToLower()){

		$recordsCorrect++
		ShowSuccess "The UPN for $displayName is correct. No change required."
        
        if($CsvLog){
		$datastring = $displayName + "," + $oldUPN + "," + $newUPN + "," + "Not Run" 
        }

	}
	else{
		try{
			ShowProgress "The current UPN is $OLDUpn"
			ShowProgress "Setting the UPN to $NewUPN"

			set-aduser $_ -userprincipalname $NewUPN 

		}
		catch{

			ShowError "ERROR - There was a problem updating the user $displayName."
			ShowError "ERROR details - $error"

		}

		Finally{ 
			if(!$error){

				ShowSuccess "Successfully updated the UPN for user $displayName to $NewUPN"
				$recordsUpdated++
                
                if($CsvLog){
				$datastring = $displayName + "," + $oldUPN + "," + $newUPN + "," + "Success" 
                }

			}
			else{
				ShowError "There was an error updating the UPN for user $displayName."
				ShowError "The latest error logged was - $error"
				$recordsError++

                if($CsvLog){
				$datastring = $displayName + "," + $oldUPN + "," + $newUPN + "," + "Failed" 
                }
			}
		}
	}
    
    if($CsvLog){
	Out-File -FilePath $logCSV -InputObject $datastring -Encoding UTF8 -append
    }

} # .NOTE End of ForEach AD User

ShowInfo "$recordsProcessed records processed."
ShowInfo "$recordsCorrect records didn't require an update."
ShowInfo "$recordsUpdated records were updated successfully."
ShowInfo "$recordsError records errored."

ShowProgress "------------Processing Ended---------------------"


$end = Get-Date;
ShowProgress "Script ended at $end";
$diff = New-TimeSpan -Start $start -End $end
ShowProgress "Time taken $($diff.Hours)h : $($diff.Minutes)m : $($diff.Seconds)s ";

############################################################################
# Script end   
############################################################################