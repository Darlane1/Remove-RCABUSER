<#
###### Offboard-rcabuser V1.4 ######
.EXAMPLE
	remove-RCABUser <user>

.DESCRIPTION 
	Offboard users from AD and other services.
	
.INPUTS 
	Ticket Number
    SamAccountName of user

.NOTES 
	Author: Darlane Tang
	Version: 1.4

.Issues/Errors
	Will not display the number at the beginning of the script if the user has been migrated to Teams
	Will not display a name if the user does not have a first and last name set in AD 
#>
#----------------------------------------------------------------------------------------------------------
#
#                                          Function Definition
#
#--------------------------------------------------------------------------------------------------------
$JsonData = Get-content -raw -path '.\Remove-RCABUSER\data.json'
#Converts data to readible powershell info
$JsonInfo = $JsonData | ConvertFrom-Json

function remove-RCABUser {
    [CmdletBinding()]
    Param
    (
        #The username of the user that will be offboarded
        [Parameter(Mandatory=$true,
                   Position=0)]
        [string]$user
    )
	
$ErrorActionPreference = 'Stop'
$GetUser = Get-AdUser -Filter "SamAccountName -eq '$user'" -Properties *
$OU = $GetUser.DistinguishedName
$GetUserEmail = $GetUser.UserPrincipalName
$GetUserGroups = Get-AdPrincipalGroupMembership -Identity $user | Where-Object {$_.name -ne "Domain Users"}
$GetSFBUser = Get-CsUser | Where-Object {$_.identity -eq $OU}
$GetUserEx = Get-Mailbox -Filter "SamAccountName -eq '$user'"
$GetUser365 = Get-RemoteMailbox -Filter "SamAccountName -eq '$user'"
$Month = (get-date).ToString('MMMM')
$Today = Get-Date -Format "dd/MM/yyyy HH:mm"
$DeletedOU = $JsonInfo.OU[0].Pathway

#Flagged groups
$Flaggedgroups = $($JsonInfo.Flaggedgroups[0].Name), $($JsonInfo.Flaggedgroups[1].Name)

$ConnectOnPrem = $JsonInfo.Scripts[0].Pathway
$RecentOffboards = $JsonInfo.Logs[1].Pathway

function CallOffboard {
	#Will try and reset the password 3 times
	$Attempts = 1
	$MaxAttempts = 3
	$PasswordReset = $false
	$Pass = ConvertTo-SecureString (([char[]]([char]33..[char]47) + [char[]]([char]65..[char]90) + ([char[]]([char]97..[char]122)) + 0..9 | Sort-Object {Get-Random})[0..19] -join '') -AsPlainText -Force    
	do{
		try{
			Set-ADAccountPassword -Identity $user -NewPassword $Pass -Reset -ErrorAction stop
			$PasswordReset = $true
		}
		catch{
			Write-Warning "There was an issue resetting the password. Trying again..."
		}
		$Attempts++
	}
	until($Attempts -gt $MaxAttempts -or $PasswordReset)
	if($PasswordReset -eq $true){
		Write-Host "Password Reset"
	}
	else{
		Write-Warning "Unable to reset password after $($MaxAttempts) attempts"
	}
	
    #Clears the AD number
    if($GetUser.telephonenumber){
        Write-Host "The number displayed on their AD was $($GetUser.telephonenumber). Cleared!"
        Set-ADUser -Identity $user -Clear telephonenumber
    }
    else {
        Write-Host "No number in AD to clear"
        }

    #Clears the AD manager
    if($GetUser.manager){
        $Manager = Get-AdUser $GetUser.manager -Properties *
        Write-Host "Their manager was $($Manager.GivenName) $($Manager.Surname). Cleared!"
        Set-ADUser -Identity $user -Clear manager
    }
    else {
        Write-Host "No manager detected"
        }
		
	#Hide user from the GAL
	try{
		Set-ADUser -Identity $GetUser -Replace @{msExchHideFromAddressLists=$True} -ErrorAction stop
		Write-Host "Hiding from the GAL"
	}
	catch{
		Write-Warning "Was unable to hide from the GAL"
	}

    #Convert the mailbox to shared.
    if ($GetUser365) {
        try {
            Write-Host "Setting 365 mailbox to Shared"
            Connect-ExchangeOnline -Showbanner:$false
            Set-Mailbox -Identity $GetUserEmail -Type Shared -Confirm:$false -force 
            Disconnect-ExchangeOnline -Confirm:$false
            #Will call for an existing script that will reconnect with OnPrem exchange modules
            Import-PSSession $session -WarningAction SilentlyContinue | Out-Null 
			Set-RemoteMailbox -Identity $user -Type Shared | Out-Null
			Write-Host "Successfully converted mailbox to shared"
			
        }
		catch [System.Net.Http.HttpRequestException]{
        Write-Warning "There were issues converting the mailbox to shared. $($_.Exception.Message)`nPlease close the ISE and open it up again and run the script to complete converting the mailbox - or do it manually"
        }
        catch {
            Write-Warning "Failed to convert the mailbox to shared. $($_.Exception.Message)"
        }
    }
	#Converting a Onprem mailbox to shared
    elseif ($GetUserEx){
        try {
            Write-Host "Could not find in mailbox on 365"
            Set-Mailbox -Identity $GetUserEmail -Type Shared -ErrorAction stop
            Write-Host "Converted OnPrem mailbox to shared successfully"
        }
        #If 365 and Onprem weren't found
        catch {
            Write-Warning "Failed to convert the mailbox to shared."
        }
    }
    else {
        Write-Warning "An error occurred when trying to convert the mailbox or mailbox doesn't exist."
    }

    #Get users info/notes and replace it with current date
    try {
        $GetNotes = $GetUser.info
        $NewNote = "$GetNotes `r`nOffboarded $($Today)"
        Set-Aduser $GetUser -Replace @{info = $NewNote} -ErrorAction stop 
        Write-Host "Notes left in AD"
    }
    catch {
        $Error[0].Exception.Message
    }
	
	#Find if user is in any of the flagged groups
	foreach($usergroup in $GetUserGroups){
		foreach($flaggroup in $Flaggedgroups){
			if($usergroup.name -like $flaggroup){
				write-warning "User is in $($flaggroup). Please manually remove them from this service"
			}
		}
	}

    #Remove security groups
    Write-Host 'Removing from the following groups:'
        foreach($group in $GetUserGroups){
            Write-Host $group.name
            Remove-AdPrincipalGroupMembership -Identity $user -memberof $GetUserGroups -Confirm:$false
        }
		
    #Remove SFB account if exists
    if($GetSFBUser){
		try{
			Disable-CsUser -Identity $GetUserEmail -ErrorAction stop
            Write-Host "Disabling SFB account"
            Write-Warning "Ensure the number - $($GetSFBUser.LineURI) is free in smartsheets" 
		}
		catch{
			Write-Warning "Issue removing the user from SFB"
			Write-Host "An error occurred while removing user from SFB: $($_.Exception.Message)"
		}
    }
    else{
        Write-Host "No SFB account was found to disable"
    }
	
	$Successfulexport = $false
	do{
		try{
			$GetUser | Select-Object @{Name="Date";Expression={Get-Date -Format "dd/MM/yyyy"}},@{Name="SAMAccountName";Expression={$user}}, GivenName, Surname, SID | Export-Csv -Path $RecentOffboards -NoTypeInformation -Append -ErrorAction stop -force
			Write-Host "Added name to CSV"
			$Successfulexport = $true
		}
		catch [System.IO.IOException]{
			Write-warning "The excel file at $($RecentOffboards) is already open. Please close it to continue"
			$OpenFile = Read-Host -Prompt "Click 'Y' once you have closed the excel file. Press 'S' to Skip" 
			if ($OpenFile -eq 'y'){
				Write-Host "Trying again..."
			}
			elseif ($OpenFile -eq 's'){
				break
			}
		}
		catch{
			Write-Warning "Was unable to export the name to the csv file"
			$Error[0].Exception.Message
			break
		}
	}
	until ($successfulexport -eq $true)
	
	#Disables the user
	Disable-AdAccount $user
	Write-Host "Disabling $($GetUser.SamAccountName)"
	
	Write-Host "Users current OU is $OU"
	$GetUser | Move-AdObject -TargetPath $DeletedOU
	Write-Host "Moving user to 'CC Removed\$($Month)'"
	
    Write-Host "Emailing HR offboarding details. To cancel, press CTRL + C"
    Start-Sleep 3 
    EmailHRIS
    Write-output "Emailed HR"

    #Should handle any terminating errors like CTRL+C and do the following
    trap{
        Write-Output "Script ended prematurely"
        Stop-Transcript -ErrorAction SilentlyContinue
        exit
    }
}

function Createtranscript{
    [string]$TicketName = read-host -prompt "Ticket Number" #Prompt what the ticket number is
    $LogPath = $JsonInfo.Logs[0].Pathway
    $SaveName = "$LogPath\$($TicketName)_$($user).txt"
    if (-not(Test-Path $LogPath)){
        New-Item -Path $LogPath -ItemType Directory
        Start-Transcript -Path "$SaveName" -Append
    }
    else{
        Start-Transcript -Path "$SaveName" -Append
    }
}

$PSDefaultParameterValues = @{
    "Send-MailMessage:from"="Service Desk <$($JsonInfo.Contacts[2].Email)>";
    "Send-MailMessage:smtpServer"="$($JsonInfo.SMTP.Address)"
}
function EmailHRIS{
    [string]$SendToAddress = $($JsonInfo.Contacts[0].Email)
    $Subject = "User offboarded: $($GetUser.GivenName) $($GetUser.Surname)"

    $Body = "
        <font face=""verdana"">
        <p>Hello team,</p>
        <p>We have received an offboarding request for the following user. They have been offboarded</p>
        <p><b>Name:</b> $($GetUser.GivenName) $($GetUser.Surname)</p>
        <p><b>Employee Number:</b> $($GetUser.EmployeeNumber)</p>
        <P>If this was a mistake, please let us know by responding to this email.</p>
        <p>Many thanks. <br></p>
        <a href=""mailto:$($JsonInfo.Contacts[2].Email)""?Subject=RE:$subject"">Service Desk</a>
        </font>"
    Send-Mailmessage -to $SendToAddress -subject $subject -body $body -bodyasHTML -priority Normal 
}
#----------------------------------------------------------------------------------------------------------
#
#                                          MAIN EXECUTION
#
#--------------------------------------------------------------------------------------------------------
##AUTHENTICATION for connecting to exchange
$maxAttempts = 0
$authenticated = $false
[string]$logonuseremail = ([adsisearcher]"(samaccountname=$env:USERNAME)").FindOne().Properties.mail
do {
    if($credentials){
        try{
            $session = New-PSSession -ConfigurationName microsoft.exchange -Credential $credentials -ConnectionUri $JsonInfo.Connection.URI -ErrorAction Stop
            $authenticated = $true
        }
        catch{
            Write-Warning "Incorrect Password"
            $global:credentials = Get-Credential -Credential $logonuseremail -ErrorAction Stop
            $maxAttempts++
        }
    }
    else{
        try {
            $global:credentials = Get-Credential -Credential $logonuseremail -ErrorAction Stop
            $session = New-PSSession -ConfigurationName microsoft.exchange -Credential $credentials -ConnectionUri $JsonInfo.Connection.URI -ErrorAction Stop
            $authenticated = $true
        } 
    catch {
            Write-Warning "Incorrect Password"
            $maxAttempts++
        }   
    }
} until ($authenticated -or $maxAttempts -eq 3)

if ($authenticated) {
    try {
        Import-PSSession $session -WarningAction SilentlyContinue -ErrorAction SilentlyContinue -ErrorAction Stop
        Write-Host "Imported OnPrem"
    } catch {
        Write-Host "OnPrem Modules already added"
    }
}

$validresponse = @("y", "Y", "n", "N")
Createtranscript | Out-Null
if ($GetUser){
    Write-Host "You are about to offboard the user: " -NoNewline
    Write-Host "$($GetUser.GivenName) $($GetUser.Surname)" -BackgroundColor Red
    Write-Host "Employee Number: " -NoNewline
    Write-Host "$($Getuser.EmployeeNumber)" -BackgroundColor Red
    if($GetSFBUser){
        Write-Host "This user has a SFB account/number. Their number is $($GetSFBUser.LineURI)."
    }
    Write-Host "Before you remove them, double-check their last day on the ticket"
    do {
        $response = Read-Host -Prompt "Do you want to continue? (Y/N)"
    } until (
        $validresponse.Contains($response)
    )
        if($response -eq "y"){
            do {
                $mailrequired = read-host -prompt "Does anybody need access to the mailbox? (Y/N)"
                  } until(
                        $validresponse.Contains($mailrequired)
                    )
              if($mailrequired -eq "y"){
                do {
                    $accessprompt = Read-Host -Prompt "Provide the username of the user that require full access to the mailbox. Type 'S' to skip"
                    if ($accessprompt -eq "S") {
                        Write-Host "Skipping step"
                        break
                    }
                    if ($accessprompt -match "^[a-zA-Z-]+$"){
                        try {
                            $giveaccessuser = Get-AdUser $accessprompt -Properties * -ErrorAction Stop
                        }
                        catch {
                            Write-Warning "Could not find anyone with this name: $($accessprompt) `nPlease try again"
                        }
                        if ($giveaccessuser) {
                            Write-Host "Trying to give " -NoNewline 
                            Write-Host "$($giveaccessuser.UserPrincipalName) " -ForegroundColor Green -NoNewline
                            Write-Host "full access to " -NoNewline
                            Write-Host "$($GetUserEmail)..." -ForegroundColor Green
                            try {
                                Add-Mailboxpermission -identity $GetUserEmail -user $giveaccessuser.UserPrincipalName -AccessRights fullaccess -ErrorAction Stop | Out-Null
                                break
                            }
                            catch {
                                If($GetUser365){
                                    Write-Host "This is a 365 user. Establishing a 365 session" -ForegroundColor DarkGray
                                    Connect-ExchangeOnline -Showbanner:$false
                                    Add-MailboxPermission -Identity $GetUserEmail -User $giveaccessuser.UserPrincipalName -AccessRights fullaccess -ErrorAction Stop | Out-Null
									Write-Host "Access provided" -ForegroundColor Green
                                    Disconnect-ExchangeOnline -Confirm:$false | Out-Null
                                    Import-PSSession $session -WarningAction SilentlyContinue | Out-Null 
                                }
                                else {
                                Write-Warning "Wasn't able to give access! Mailboxes may not be on the same environment or 365 session isn't established`nPlease try again"
                                $Error[0].Exception.Message
                                }
                            }
                        }
                    }
                    else {
                        #If the input does not meet the regex, it will go back to "Provide the username..."
                        Write-Host "Invalid Input. Please try again"
                    }
                } until (
                    $false
                )
                  CallOffboard
                }             
              else{
                CallOffboard 
              }
        }
        #enter 'n' if you do not wish to continue with the offboarding
        elseif($response -eq "n"){
            Write-Host "Exiting..."
        }
}
#If the user does not exist
else{
    Write-Host "This user does not exist"
}
Stop-Transcript
}