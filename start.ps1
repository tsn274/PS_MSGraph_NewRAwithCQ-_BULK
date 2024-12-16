# PS connection with scopes
#connect-azuread
#Connect-MgGraph -Scopes "User.ReadWrite.All, Directory.ReadWrite.All,Groupmember.ReadWrite.All"
#Connect-MicrosoftTeams


## logfile.
Start-Transcript

<#
Prepare your CSV file: Create a CSV file with the necessary columns c:\temp\list.csv:

Naam,Phonenumber,Agents
Sales Queue,+3120598xxxx,user1@domain.com;user2@domain.com
Support Queue,+3120598xxxxa,user3@domain.com;user4@domain.com
Facilities Queue,+3120598xxxxb,user5@domain.com;user6@domain.com
#>


#
# Welke Faculteit dienst
$department = Read-Host "Geef dienst/faculteit (bij. IT)"
#defaul instellingen voor Callqueueu
$CQAgentsRoutingmethod = "Longestidle"
$cqpresencerouting = $true
$cqagentsoptout = $true

#import file volgens format
$importlist = Import-Csv -Path "C:\temp\list.csv"

foreach ($row in $importlist) {

    $rowNaam = $row.Naam
    $rowPhonenumber = $row.Phonenumber
    $rowAgents = $row.Agents
    $displayname = $department +" " + $rowNaam

    $words = $displayname -split " "

    $finaldisplayname = $words[0] + "-" + ($words[1..($words.Length - 1)] -join "_")


    $domain = "@vunl.onmicrosoft.com" ##DEZE AANPASSEN VOOR PRODUCTIE OF TEST

    $upn = "AAD-CQ-" + $finaldisplayname + $domain
    $mailnickname = $finaldisplayname -replace "_","-"  
   
    Write-Host "
    ----------------------
    CallQueue Naam: $displayname
    UPN ResourceAccount: $upn
    Faculteit/Dienst: $department
    PhoneNumber Resource Account: $rowphonenumber
    CallQueue Routing Methode: $CQAgentsRoutingmethod
    Agents in CallQueue: $rowAgents
    CallQueue Presence-based routing AAN: $cqpresenceRouting
    CallQueue Agents Opt-out AAN: $cqagentsoptout
    CallQueue MailNicknam: $mailnickname
    "
    # add date and time
    Write-Output (Get-Date -Format "yyyy-MM-dd HH:mm:ss") 

    #check if ra exist
    Write-Output "Check if RA exist..." 

$RAExist = Get-CsOnlineApplicationInstance $upn -ErrorAction SilentlyContinue

if (-not $RAExist) {
    Write-Host $summary -ForegroundColor Cyan
    Write-Output "Continue creating RA" 
    $confirmation = Read-Host "Are you sure? (yes to continue, no to stop)"
} else {
    Write-Host "RA CQ already exists..." -ForegroundColor Red
    exit
}

if ($confirmation -eq "yes") {
    Write-Output "Continuing..." 
    $resourceAccountParams = @{
        UserPrincipalName = $upn
        ApplicationId     = "11cd3e2e-fccb-42ad-ad00-878b93575e07"
        DisplayName       = $displayname
    }

    try {
        Write-Output "Make new ResourceAccount..." 
        # Uncomment the line below to actually create the resource account
        # $resourceAccount = New-CsOnlineApplicationInstance @resourceAccountParams
        Write-Output "Resource account created successfully." 
    } catch {
        Write-Error "An error occurred: $_"
    }

    # Assign License
    $entitlementgroupParams = @{
        GroupID = '4b35443a-4385-4a6b-82f0-7869b015db92'
        DirectoryObjectID = $resourceAccount.ObjectId
    }

    try {
        Write-Output "Assign License to ResourceAccount..." 
        # Uncomment the line below to actually assign the license
        # $entitlement = New-MgGroupMember @entitlementgroupParams
        Write-Output "Assign License successfully" 
    } catch {
        Write-Error "An error occurred: $_"
    }

    # Pause for license processing
    Write-Host "10 sec wachten..."
    Start-Sleep 10

    # Assign Phone Number
    $phonenumberassignParams = @{
        Identity        = $resourceAccountParams.UserPrincipalName
        PhoneNumber     = $rowPhonenumber
        PhoneNumberType = 'DirectRouting'
    }

    try {
        Write-Output "Assign Phonenumber to ResourceAccount"
        # Uncomment the line below to actually assign the phone number
        # $phonenumberassign = Set-CsPhoneNumberAssignment @phonenumberassignParams
        Write-Output "Phonenumber $phonenumberassign assigned to ResourceAccount..." 
    } catch {
        Write-Error "An error occurred: $_"
    }


# parameters callqueue
$callQueueParams = @{
    Name = $displayname
    Conferencemode = $true
    RoutingMethod = $CQAgentsRoutingmethod 
    PresenceBasedRouting = $cqpresencerouting
    UseDefaultMusicOnHold = $true
    #OboResourceAccountIds = $resourceAccount.Id  ########## caller id meegeven aan callqueue CHECKEN OF DIT WERKT.
    #RoutingMethod = LongestIdle
    #AgentAlertTime = 30
    AllowOptOut = $cqagentsoptout
    #DistributionLists = @("support@yourdomain.com")
    #WelcomeMusicAudioFileID = "audioFileSupportGreetingID"
    #MusicOnHoldAudioFileID = "audioFileSupportHoldInQueueMusicID"
    #OverflowAction = "SharedVoicemail"
    #OverflowActionTarget = "support@yourdomain.com"
    OverflowThreshold = '50'
    OverflowAction = "DisconnectWithBusy"
    #OverflowSharedVoicemailAudioFilePrompt = "audioFileSupportSharedVoicemailGreetingID"
    #EnableOverflowSharedVoicemailTranscription = $true
    TimeoutAction = "Disconnect"
    #TimeoutActionTarget = "support@yourdomain.com"
    TimeoutThreshold = '1200'
    #TimeoutSharedVoicemailTextToSpeechPrompt = "We're sorry to have kept you waiting and are now transferring your call to voicemail."
}

try {
    # Create the new Call Queue
    Write-Output "Creating new Call Queue..." 
    # Uncomment the line below to actually create the callqueue
    #$newCallQueue = New-CsCallQueue @callQueueParams -WarningAction SilentlyContinue

    # Check if the Call Queue was created successfully
    if ($newCallQueue) {
        Write-Output "Call Queue created successfully." 
    } else {
        Write-Output "Failed to create Call Queue." 
    }
} catch {
    # Handle any errors that occur during the creation process
    Write-Output "An error occurred: $_" 
}


#### vanaf hier verder met agents toevoegen

 # agents toevoegen aan call queue
    ## ONDERSTAANDE IS NIET ZO BEST KAN MAKKELIJKER
   # Initialize arrays to store user GUIDs, emails, and not added users
$userGuids = @()
$userlist = @()
$notAddedUsers = @()

# Loop through each agent in the row.Agents, split by semicolon
foreach ($item in $row.Agents.Split(';')) {
    # Get user details using Microsoft Graph API
    $useremail = Get-MgUser -UserId $item -ErrorAction SilentlyContinue
    if ($useremail) {
        # Add user ID to the GUIDs array
        $userGuids += $useremail.Id
        # Add user email to the user list array
        $userlist += $useremail.Mail
    } else {
        # Add user to the not added users array if not found
        $notAddedUsers += $item
    }
}

try {
    Write-Output "Adding the eendjes" 
    # Set the call queue with the collected user GUIDs
    #Set-CsCallQueue -Identity $newCallQueue.Identity -Users $userGuids -WarningAction SilentlyContinue > $null
    Write-Output "eendjes added... $userlist" 
}
catch {
    Write-Output "An error occurred: $_" 
}

# Output the users who were not added
if ($notAddedUsers.Count -gt 0) {
    Write-Output "The following users were not added: $notAddedUsers" 
} else {
    Write-Output "All users were added successfully." 
}

Write-Output "======================================" 

} elseif ($confirmation -eq "no") {
    Write-Output "Stopping..." 
    exit
} else {
    Write-Output "Invalid input. Please enter 'yes' or 'no'." 
}
}

#stop logging
Stop-Transcript