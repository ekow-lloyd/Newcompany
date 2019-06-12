<#
.SYNOPSIS
Create a new or modify an existing CloudCom user in the CloudCom AD.
TEST
.DESCRIPTION
Create a new or modify an existing CloudCom user in the CloudCom AD.
Two parameter sets ("Init" and "Scheduled") exist.  This is to make it easier to call the script when it's initially called (when you want to read CSV files) or when it's called as part of a scheduled task.
Only paramaters that are members of a parameter set can be called in an single instance.

This is used so the script can read the CSV file(s) (init) and process the request based on the startdate value in the read CSV file.
If the startdate of the user is within 48 hours of the scrpit run then it'll automatically add the user to AD at the time of script run.
Otherwise, if the startdate of the user is beyond 48 hours of the script run, the script will *automatically* create a scheduled tasks to add the user within 48 hours of the CSV startdate value.

.PARAMETER isScheduled
Type: SWITCH 
Mandatory: Yes (Init)
Set: Init, Scheduled
Tells the script whether or not to run in a scheduled task mode ($true) or 'input from csv' mode ($false)
SWITCH paramaters do not need values associated.  In our case, running Update-CloudComAD.ps1 -isScheduled is the same as saying Update-CloudComAD.ps1 -isScheduled $true 

.PARAMETER pFirstName
Type: String
Mandatory: Yes
Set: Scheduled

The firstname of the user.  Supplied as a parameter and value to the script when run from a scheduled task.

.PARAMETER pLastName
Type: String
Mandatory: Yes
Set: Scheduled

The last name of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER pSAM
Type: String
Mandatory: Yes
Set: Scheduled

The SamAccountName of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER pUserName
Type: String
Mandatory: Yes
Set: Scheduled

The username of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER pOU
Type: String
Mandatory: Yes
Set: Scheduled

The OU the user will belong to.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER pStartDate
Type: String
Mandatory: Yes
Set: Scheduled

The Start Date of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER pEndDate
Type: String
Mandatory: Yes
Set: Scheduled

The End Date of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER pCompany
Type: String
Mandatory: Yes
Set: Scheduled

The Company the user belongs to. Supplied as a parameter and value to the scirpt when run from a scheduled task.

.INPUTS
When run in "Init" set the path of the CSV file(s) are required.

.OUTPUTS
Outputs a transaction log to the user's Desktop ($env:username\desktop\) and writes to the local Event Viewer "Application" log under the source "Update-CloudComAD".

.EXAMPLE

.\New-CloudComUser.ps1

.NOTES
There is NO NEED for a user to use any of the paramaters that start with "p".  These are used by the script. If you want to schedule this script as a task (automated) then you just need to point the scheduled task to this script file with no paramters just like the example.

#>
#Requires -RunAsAdministrator

Param(
    [CmdletBinding(DefaultParameterSetName='Init')]
    # Whether or not this is a scheduled task...
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")] #include $isScheduled in both (scheudled and init) parameter sets.
    [Parameter(Mandatory=$false,ParameterSetName="Init")]
    [switch]
    $isScheduled,
    # Let's define all required parameters when creating a user when it's a scheduled task.  Scheduled tasks require additional parameters because the initial CSV that was loaded will no longer be used.  Instead, all values from the CSV will be stored as arguments (parameters) to the script within the scheduled task.
    # First Name
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pFirstName,
    # Last Name
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pLastName,
    # SAM
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pSAM,
    # end date
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pEndDate,
    # company
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pCompany,
    # copyuser name
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pCopyUser,
    # UPN
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pUPN,
    # Full Name
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pFullName,
    # Email Address
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $pEmail
)

$DebugPreference = "Continue" #comment this line out when you don't want to enable the debugging output.
$WarningPreference = "Continue" #comment this out when testing is completed.

Write-Debug "Current parameter set: $($PSCmdlet.ParameterSetName)"
#$LogFolder = "$env:userprofile\desktop\logs\" #log file location.
$TranscriptLog = "$($env:userprofile)\desktop\logs\transcript.log"
Start-Transcript -Path $TranscriptLog -Force
$csvPath = "C:\testfc\" #changeme - Location where the website is delivering the CVS files.  Only a directory path is needed, do not enter a full path to a specific CSV file.
$ScriptFullName = -join($PSScriptRoot,"\$($MyInvocation.MyCommand.Name)") #Dynamically create this script's path and name for use later in scheduled task creation.
$Global:ConfigLocation = "$env:appdata\Update-CloudComAD.config" #DO NOT MODIFY UNLESS YOU UNDERSTAND HOW THE MS DPAPI WORKS.
$Global:sendmail = $true #if you want the script to send email notifications set to $true, else set to $false

function Format-CsvValue {
    [CmdletBinding()]
    param (
        #Sets whether or not we want to format the provided string into 'title' (aka Proper) case when using named values.
        #When isTitleCase = $true the function will take the input string ($sValue) and format it to proper(title) case and will also remove leading and trailing whitespaces.  Example; "JoHN SmITH" will return "John Smith" or "   JaNE " will return "Jane" (removed whitespaces and set to title case).
        [Parameter(Mandatory = $false)]
        [bool]
        $isTitleCase = $false,
        #The string value that's passed into the function to properly format.
        #Example: Format-CsvValue -isTitleCase $true -sValue $mvar
        #Example: To only remove whitespace from a string-> Format-CsvValue -sValue $myvar
        [Parameter(Mandatory = $true)]
        [string]
        $sValue
    
    ) #=>Params
  
    begin {
        #no variables or intitializations to declare.
    } #=>begin
  
    process {
        if ($isTitleCase) {
            #isTitleCase is set to true so let's format it...
            $rValue = $((Get-Culture).TextInfo.ToTitleCase($sValue.ToLower())).Trim() #trim leading/trailing whitespace AND convert to title case string format.
        }
        else {
            #only whitespace trim is required, process that request.
            $rValue = $sValue.Trim() #Remove leading/trailing whitespace.
        }#=>if/isTitleCase
    }#=>process
  
    end {
        #return the value through the function.
        $rValue
    }
} #=>Format-CsvValue

Function Write-CustomEventLog {
    [CmdletBinding()]
    param(
        # What message to write to the event viewer.
        [Parameter(Mandatory=$true)]
        [string]
        $message,
        # Type
        [Parameter(Mandatory=$true)]
        [ValidateSet('Information','Warning','Error')]
        [string]
        $entryType
    )

    Begin {
        $eventSourceExists = [System.Diagnostics.EventLog]::SourceExists("Update-CloudComAD")
        if(-not($eventSourceExists)) {
            try {
                New-EventLog -LogName Application -Source 'Update-CloudComAD'
            }
            catch {
                Write-Debug 'Unable to create new application source.'
            }
        }#=>if not $eventSourceExists
    }#=>Begin

    Process {
        switch ($entryType) {
            "Information" { [int]$EventID = 1000 }
            "Warning" { [int]$EventID = 2000 }
            "Error" { [int]$EventID = 3000}
        }
        Write-EventLog -LogName Application -Source 'Update-CloudComAD' -EntryType $entryType -EventId $EventID -Message $message
    }
}
Function Add-ConfigFile {
    <#
    .SYNOPSIS
        Create a custom config file.
    .DESCRIPTION
        Create a custom config file to *securely* store credential objects on the local machine using Microsoft's DPAPI.
        Microsoft's DPAPI relies on built-in encryption/decryption keys based on the USER that has run the script.
        When the config file is created it stores a file in the user's $env:appdata folder named Update-CloudComAD.config - only that user can decrypt the securestrings stored there.
    #>
}

Function Send-CustomMail{
    <#
    .SYNOPSIS
        Sends an email to specified users/groups
    .DESCRIPTION
        If enabled this function will send an email notification containing the results of this script.
    .PARAMETER setMailType
        [REQUIRED][STRING]
        Sets the email subject based on the type of message being sent.  Options are Success, Information, Warning, Error
    .PARAMETER prependBody
        [OPTIONAL][STRING]
        Allows you to customize the body.  By default the body of the message is only specific notices of errors or warnings that the script produces and may include the transcipt log.
    .PARAMETER appendSubject
        [OPTIONAL][STRING]
        Allows you to customize the subject.  By default the subject of the message is created by the script - adding this parameter and the string will add the string to the end of the pre-defined subject line.
    .INPUTS
        Reads the transaction log file and imports it into the body of the message.
    .OUTPUTS
        Sends an email message.
    .NOTES
        You will also need to fill in the variables for $mailProperties in the "Begin" block of this function.
    #>

    [CmdletBinding()]
    Param(
        #Sets the email subject line based on the type specified below.
        [Parameter(Mandatory=$true)]
        [ValidateSet("Success","Information","Warning","Error")]
        [string]
        $setMailType,
        #Allows custom text to the body of the email message.  This string will prepend (be at the beginning of) the message body.
        [Parameter(Mandatory=$false)]
        [string]
        $prependBody,
        #Allows custom text to the subject of the email message.  This string will append (be at the end of) the message subject.
        [Parameter(Mandatory=$false)]
        [string]
        $appendSubject
    )

    Begin {
        if($Global:sendmail) {
            #$sendmail is set to true so we'll continue with the processing of the script.
            $anonymousSMTP = $false #CHANGEME - set to $true if your SMTP server allows anonymous authentication.  Else, you need to set to $false and configure the $SMTPUser and $SMTPPass variables below.
            $mailProperties = @{
                SmtpServer = '' #CHANGEME - hostname or IP of the SMTP server.
                Port = '25' #CHANGEME - defaults to port 25 for SMTP.  Will need to be changed if you use SMTPS
                UseSsl = '' #CHANGEME - set to $true if your SMTP server uses SMTPS. Don't forget to change the port # as well. 
                From = '' #CHANGME - the email address in which the message will be sent from.
                To = '' #CHANGEME - the email address(es) in which to send the message.
                Subject = '' #DO NOT MODIFY
                Body = '' #DO NOT MODIFY
            }

            if (-not($anonymousSMTP)) {
                #SMTP requires auth so we configure the credentials.
                $SMTPUser = '' #CHANGEME - only change me if your SMTP server requires authentication - $anonymousSMTP = $false
                $SMTPPass = ConvertTo-SecureString 'SuperSecretPassword' -AsPlainText -Force #CHANGEME - only change 'SuperSecretPassword' if your SMTP server requires authentication.
                $SMTPCreds = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SMTPUser, $SMTPPass #DO NOT MODIFY
                $mailProperties['Credential'] = $SMTPCreds
            }#=>if -not $anonymousSMTP

            #create our default subject lines.
            switch ($setMailType) {
                "Success" {
                    $mailProperties['Subject'] = "[SUCCESS] Update-CloudComAD Script Run $($appendSubject)"
                }
                "Information" {
                    $mailProperties['Subject'] = "[INFORMATION] Update-CloudComAD Script Run $($appendSubject)"
                }
                "Warning" {
                    $mailProperties['Subject'] = "[WARNING] Update-CloudComAD Script Run $($appendSubject)"
                }
                "Error" {
                    $mailProperties['Subject'] = "[ERROR] Update-CloudComAD Script Run $($appendSubject)"
                }
            }

            #create our deafult body.
            $getTranscript = Get-Content -Path $Global:ConfigLocation 
            $mailProperties['Body'] = "$($prependBody) `n`n Please review the details of the script run below. `n`n $($getTranscript)"

        } else {
            #$sendmail is set to false so we'll exit the function and not send an email.
            Write-Debug "`$sendMail is set to false. Not sending any email notifications."
            return
        }#=>else $global:sendmail
    }

    Process {
        #Time to actually send the mail.
        try {
            Send-MailMessage @mailProperties -ErrorAction 'Stop'
        }
        catch {
            Throw $_.Exception.Message
        }
        Return "Mail message sent."
    }
}

Write-Debug "Checking to see if ActiveDirectory PS module is imported."
If(-not(Get-Module ActiveDirectory)) {
    Write-Debug "ActiveDirectory PS Module not imported. Importing."
    Import-Module ActiveDirectory
} else {
    Write-Debug "ActiveDirectory PS Module *is* imported."
}

if (!($isScheduled)) {
    Write-Debug "This is not a scheduled task so we can safely assume this is an initial read of a CSV file. Looking for all CSV files in $($csvPath) that are NOT readonly."
    #since we are anticipating *dynamically* named CSV files let's find all CSV files we have yet to process.
    $csvFiles = Get-ChildItem -Path $csvPath -Filter "*.csv" -Attributes !readonly+!directory
    if ($csvFiles) {
        $csvCount = ($csvFiles | Measure-Object).Count
        Write-Debug "Found $($csvCount) unprocessed CSV file(s): $($csvFiles)"
        foreach ($csvFile in $csvFiles) {
            Write-Debug "Processing CSV file $($csvFile.FullName)"
            try {
                $Users = Import-CSV $csvFile.FullName
            }
            catch {
                #We need to check if the csvFiles count is greater than 1. If it is, we can move to the next file. If it's not, we need to throw an error and exit this script.
                if ($csvCount -gt '1') {
                    Write-CustomEventLog -message "Unable to import CSV file: $($csvFile.FullName). This is a fatal error for this csv file. Continuing to next file. Error message is: `n`n $($_.Exception.Message)" -entryType "Warning"
                    Write-Debug "Unable to import our CSV file: $($csvFile.FullName). This is a fatal error for this CSV file.  Continuing to next file. Error message is: $($_.Exception.Message)"
                    Continue
                } else {
                    Write-CustomEventLog -message "Unable to import CSV file: $($csvFile.FullName). This is a fatal error for this csv file and this script. Exiting script. Error message is: `n`n $($_.Exception.Message)" -entryType "Error"
                    Write-Debug "Unable to import our CSV file: $($csvFile.FullName). This is a fatal error for this CSV file and this script. Exiting script. Error message is: $($_.Exception.Message)"
                    Continue
                }
            }#=> try $Users
        
            #imported our CSV file properly.  Let's process the file for new users...
            ForEach ($User in $Users){

                #the email address in the CSV file is considered the 'primary key' when doing AD lookups.  We should check the email address in the CSV to make sure it's at least filled out. If not, we'll error out the script.
                if([string]::IsNullOrEmpty($user.Email)){
                    Write-Debug "The CSV file $($csvFile.FullName) contains an empty email address for user $($user.FirstName) $($user.LastName) we need to skip this user and subsequently this CSV file."
                    Write-CustomEventLog -message "The CSV file $($csvFile.FullName) contains an empty email address for user $($user.FirstName) $($user.LastName) we need to skip this user and subsequently this CSV file." -entryType "Warning"
                    continue
                }

                #debugging purposes...
                Write-Debug "Found the following information in the CSV File: `n`n First Name (CSV): $($User.Firstname) `n`n Last Name (CSV): $($User.Lastname) `n`n StartDate (CSV): $($User.startdate) `n`n End Date (CSV): $($User.enddate) `n`n Company (CSV): $($User.Company) `n`n Email (CSV): $($user.Email)"
                #=>debugging purposes.
                
                #We should really clear all variables in the loop to make sure they get the new information on the next loop in ForEach ($user in $users)
                $myvars = "FirstName","LastName","Email","StartDate","EndDate","Company","FullName","SAM","Username","DNSroot","UPN","oStartDate","oEndDate","templateUser","copyUser","OU","newUserAD","newUserExch","copyMailProps","oNewUserExch","setUserADProps","oChangeUserAD","changeUserAD"
                Remove-Variable -Name $myvars -ErrorAction 'SilentlyContinue'
                #=>clear variables

                #Let's properly format all the values in this *ROW* of the CSV. Trim() where necessary and change to Title Case where necessary - also create a new variable so we can use it later when creating the user in AD using the New-ADuser cmdlet.
                $FirstName = Format-CsvValue -isTitleCase $true -sValue $User.FirstName #trim and title case
                $LastName = Format-CsvValue -isTitleCase $true -sValue $User.LastName #trim and title case.
                $Email = Format-CsvValue -sValue $User.Email #trim only.
                $StartDate = Format-CsvValue -sValue $User.startdate #trim only.
                $EndDate = Format-CsvValue -sValue $User.enddate #trim only.
                $Company = Format-CsvValue -sValue $User.company #trim only since company names are rather specific on how they're spelled out.

                if ($csvFile.Name -like "NU*") {
                    #This csvFile that we're working on seems to be a New User request as defined by the "NU" in the CSV file name so we add more details.
                    $copyUser = -join(($User.copyuser).Trim()," ", ($User.copyuserLN).Trim()) #We need the fullname of the user we're copying from.
                }
                #=> End of CSV values.

                #Let's build other necessary variables that we want to use as parameters for the New-ADuser cmdlet out of the information provided by the CSV file or other sources...
                $FullName = -join($($FirstName)," ",$($LastName)) #join $Firstname and $Lastname and a space to get FullName
                $SAM = ( -join ($FirstName,".","$LastName","-$($Company -replace '\s','')")).ToLower()
                $Username = (-join($FirstName,".",$LastName)).ToLower() #this assumes that your usernames have a naming convention of firstname.lastname and makes everything lowercase.
                
                $UPN = $Email
                $Password = (ConvertTo-SecureString -AsPlainText 'Cloudcom.1' -Force)
                $oStartDate = [datetime]::ParseExact(($User.StartDate).Trim(), "dd/MM/yyyy", $null) #This converts the CSV "startdate" field from a string to a datetime object so we can use it for comparison.
                $oEndDate = [datetime]::ParseExact(($User.EndDate).Trim(), "dd/MM/yyyy", $null) #This conerts to CSV 'EndDate' field from a string to a datetime object which is required for the New-AdUser cmdlet 'AccountExpirationDate' parameter.

                #debugging purposes...
                Write-Debug "Script created these properties: `n`n `$FirstName:  $($FirstName) `n`n `$LastName: $($LastName) `n`n `$Email: $($Email) `n`n `$StartDate: $($StartDate) `n`n `$EndDate: $($EndDate) `n`n `$copyUser: $($copyUser) `n`n `$FullName: $($FullName) `n`n `$SAM: $($SAM) `n`n `$Username: $($Username) `n`n `$UPN: $($UPN) `n`n `$oStartDate: $($oStartDate)"
                #=>debugging puproses

                #Now, let's check the user's startdate as listed in the CSV file.  If startdate is within 48 hours of today's (Get-Date) date we'll create the user directly in AD.  Otherwise, we'll schedule a task to create the user at a later date.
                #First, we need to check if this is a New User request, 'startdate' only applies to new users...
                if ($csvFile.name -like "NU*") {
                    if ( $(get-date) -ge ($oStartDate).AddHours(-48) ) {
                        Write-Debug "$(Get-Date) (current script run time/date) is greater than or equal to 48 hours minus employee start date: $($oStartDate).AddHours(-48)) so we are creating the user immediately."

                        #Checking to see if a user already exists in AD with the same email address...
                        try {
                            Get-ADUser -Filter {mail -eq $Email} -ErrorAction 'Stop'
                        }
                        catch {
                            Write-Debug "Unable to check if user $($FullName) already exists in AD.  Get-ADUser returned an error when doing a lookup. $($_.Exception.Message)"
                            Write-CustomEventLog -message "Unable o check if user $($FullName) already exists in AD.  Get-ADUser returned an error when doing a lookup. $($_.Exception.Message)" -entryType "Warning"
                            continue
                        }
                        if (Get-AdUser -Filter {mail -eq $Email}) {
                            Write-Debug "A user with email address $($email) already exists in AD.  We cannot add this user."
                            Write-CustomEventLog -message "When attempting to create user $($FullName) [SAM: $($SAM)] we found another user that exists in AD using the same email address of $($email). We have to skip this user." -entryType "Warning"
                            Continue #go to next csv record.
                        }#=if get-aduser
                        else {
                            Write-Debug "No existing user in AD with email address $($email) so we can create our user."

                            $newUserAD = @{
                                'Company'                   = $Company
                                'AccountExpirationDate'     = $oEndDate
                                'ChangePasswordAtLogon'     = $true
                                'Enabled'                   = $true
                                'PasswordNeverExpires'      = $false
                            }#=>$newUserAD

                            $newUserExch = @{
                                'SamAccountName'            = $SAM
                                'UserPrincipalName'         = $UPN
                                'Name'                      = $FullName
                                'PrimarySmtpAddress'        = $Email
                                'FirstName'                 = $FirstName
                                'LastName'                  = $LastName
                                'Password'                  = $Password
                            }

                            Write-Debug "Running Get-ADUser against $($copyUser) so we can set `$templateUser to copy their properties."
                            try {
                                
                                $templateUser = Get-ADUser -filter {name -eq $copyUser} -Properties MemberOf,EmailAddress -ErrorAction 'Stop' -WarningAction 'Stop'    
                            }
                            catch {
                                Write-Debug "We were unable to find the template user $($copyUser) so we have to skip this new AD user and go to the next row in the CSV file."
                                #$failedUsers+= -join($Fullname,",",$SAM,",","We were unable to find the template user $($copyUser) so we have to skip creating new user $($FullName) and go to the next row in the CSV file.")
                                Write-CustomEventLog -message "We were unable to find the template user $($copyUser) when attempting to create new user $($FullName) with SAM $($SAM).  Skipping the creation of this user." -entryType "Warning"
                                continue #move to next CSV row.
                            }#=>try/catch $templateUser
                            
                            if (-not($templateUser)) {
                                Write-Debug "We were unable to find the template user $($copyUser) so we have to skip this new AD user and go to the next row in the CSV file."
                                #$failedUsers+= -join($Fullname,",",$SAM,",","We were unable to find the template user $($copyUser) so we have to skip creating new user $($FullName) and go to the next row in the CSV file.")
                                Write-CustomEventLog -message "We were unable to find the template user $($copyUser) when attempting to create new user $($FullName) with SAM $($SAM).  Skipping the creation of this user." -entryType "Warning"
                                continue #move to next CSV row.
                            } else {
                                Write-Debug "Get-ADUser against $($copyUser) success.  Our `$templateUser is now set."
                                #Let's get the OU that our template user belongs to and apply that to our new user...
                                $OU = ($templateUser.DistinguishedName).Substring(($templateUser.DistinguishedName).IndexOf(",")+1)
                                Write-Debug "Our OU for new user $($FullName) is $($OU) from copy of our template user $($copyUser) with OU of $($templateUser.DistinguishedName)"
                                #Let's update our $newUserAD properties with this OU...
                                $newUserExch['OrganizationalUnit'] = $OU
                            }#=>if/else not $templateuser

                            try {
                                Add-PSSnapin "*Exchange*"
                            }
                            catch {
                                Write-Debug "Unable to connect to Exchange PowerShell due to the following error $($_.Exception.Message).  This is likely a fatal error for the entire email portion of the script."
                                Write-CustomEventLog -message "Unable to connect to Exchange Powershell to create mailbox for user $($FullName) due to the following error: `n`n $($_.Exception.Message) `n`n This is likely a fatal error for the entire email portion of the script.  This error should be remedied or no email boxes will be created." -entryType "Error"
                                Stop-Transcript
                                exit
                            }#=>Add-PSSnapin

                            Write-Debug "Running Get-Mailbox using identity parameter of $($templateUser.EmailAddress)"
                            try {
                                
                                $copyMailProps = Get-MailBox -Identity $($templateUser.EmailAddress) -ErrorAction 'Stop' -WarningAction 'Stop' | Select-Object AddressBookPolicy,Database
                            }#=>try $copyMailProps
                            catch {
                                Write-Debug "Unable to Get-Mailbox for template user $($templateUser.EmailAddress) which means we are unable to activate $($FullName) in AD or Exchange. Continuing to next user."
                                Write-CustomEventLog -message "Unable to Get-Mailbox for template user $($templateUser.EmailAddress) which means we are unable to activate $($FullName) in AD or Exchange." -entryType "Warning"
                                Continue
                            }#=>catch $copyMailProps

                            if (-not($copyMailProps)) {
                                Write-Debug "Unable to Get-Mailbox for template user $($templateUser.EmailAddress) which means we are unable to activate $($FullName)'s Exchange Email. Continuing to next user."
                                Write-CustomEventLog -message "Unable to Get-Mailbox for template user $($templateUser.EmailAddress) which means we are unable to activate $($FullName)'s Exchange Email." -entryType "Warning"
                                Continue
                            } else {
                                #We got our $copyMailProps properties for Database and AddressBookPolicy so we'll just add that to the $newUserExch hashtable.
                                Write-Debug "`$copyMailProps has returned; $($copyMailProps | Out-String)"
                                $newUserExch['AddressBookPolicy'] = $copyMailProps.AddressBookPolicy
                                $newUserExch['Database'] = $copyMailProps.Database
                            }#=>if not $copyMailProps

                            Write-Debug "Creating user in Exchange and AD using New-Mailbox cmdlet.  Passing parameters: `n $($newUserExch | Out-String)"
                            try {
                                $oNewExchUser = New-Mailbox @newUserExch -ErrorAction 'Stop' -WarningAction 'Stop'
                            }
                            catch {
                                Write-Debug "Unable to create new user $($FullName) using New-Mailbox.  Error message: `n`n $($_.Exception.Message)"
                                Write-CustomEventLog -message "We were unable to add our new user $($FullName) to AD and Exchage.  Skipping this user.  Full error details below: `n`n $($_.Exception.Message)" -entryType "Warning"
                                continue
                            }
                            if(-not($oNewExchUser)) {
                                Write-Debug "Something went wrong with adding our new $($FullName) user to AD and Exchange. `n`n $($_.Exception.Message)"
                                Write-CustomEventLog -message "`$oNewExchUser returned false for some reason which means we were unable to add our new user $($FullName) to AD and Exchange. Skipping this user." -entryType "Warning"
                                continue
                            }
                            #Adding user went well now let's update the AD properties for this user that can't be done using the New-Mailbox cmdlet.
                            Write-Debug "We created our new user $($FullName) in AD and Exchange."
                            try {
                                Write-Debug "Getting AD User $($SAM) and setting properties..."
                                $setUserADProps = Get-ADUser -Identity $($SAM) | Set-ADUser @newUserAD -ErrorAction 'Stop' -WarningAction 'Stop' -PassThru
                            }
                            catch {
                                Write-Debug "Unable to modify AD user properties for $($FullName).  Continuing to next user."
                                Write-Debug "`$setUserADProps has produced `n`n $($setUserADProps)"
                                Write-CustomEventLog -message "We were unable to modify AD properties for user $($FullName).  Full error is `n`n $($_.Exception.Message)).`n`n User properties we want to modify are $($newUserAD | Out-String)" -entryType "Error"
                                continue
                            }#=> try/catch $setUserADProps
                            if(-not($setUserADProps)) {
                                Write-Debug "Unable to modify AD user properties for $($FullName).  Continuing to next user. Error and warnings displayed below: `n`n $SetADErr `n`n $SetADWarn"
                                Write-CustomEventLog -message "We were unable to modify AD properties for user $($FullName).  Full error is `n`n $($_.Exception.Message).`n`n User properties we want to modify are $($newUserAD | Out-String)" -entryType "Error"
                                continue
                            } else {
                                Write-Debug "Successfully created new Exchange mailbox and modified AD properties for user $($FullName)"
                                Write-CustomEventLog -message "Successfully created new AD User and Exchange Mailbox for $($FullName).  AD and Exchange Details included below; `n`n $($newUserExch | Out-String) `n`n $($newUserAD | Out-String)" -entryType "Information"
                            }
                        }#=>else get-aduser

                    #####Create a scheduled task#####        
                    } else {
                        Write-Debug "$(Get-Date) (current script run time/date) is NOT greater than or equal to 48 hours minus employee start date: $($oStartDate).AddHours(-48)) so we are scheduling a task to create the user later."
                        $taskaction = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument "-NoProfile -windowStyle Hidden -Command `"& $($ScriptFullName) -isScheduled -pSAM '$($SAM)' -pUPN '$($UPN)' -pFullName '$($FullName)' -pCompany '$($Company)' -pEmail '$($Email)' -pFirstName '$($FirstName)' -pLastName '$($LastName)' -pEndDate '$($EndDate)' -pCopyUser '$($copyuser)'`""
                        $tasktrigger = New-ScheduledTaskTrigger -Once -At ($oStartDate).AddHours(-48)
                        try {
                            $taskregister = Register-ScheduledTask -Action $taskaction -Trigger $tasktrigger -TaskName "Add AD User - $($FullName)" -Description "Automatic creation of AD User $($FullName) 48 hours prior to the user's startdate." -ErrorAction 'Stop'
                        }
                        catch {
                            Write-Warning $_
                        }
                        $findTask = Get-ScheduledTask -TaskName "Add AD User - $($FullName)"
                        if(-not($findTask)) {
                            Write-Debug "Our scheduled task | Add AD User - $($FullName) | was NOT created."
                            Write-CustomEventLog -message "We were unable to create a scheduled task to create user $($FullName) - $($SAM) on $($StartDate)." -entryType "Warning"
                        } else {
                            Write-Debug "Our scheduled task | Add AD User - $($FullName) | was created."
                            Write-CustomEventLog -message "Created a scheduled task to create AD User $($FullName) $($SAM) on $(($oStartDate).AddHours(-48))" -entryType "Information"
                        } #=>if/else not $findTask
                    }#=>if/get-date -ge startdate-48
                    #####/Create a scheduled task#####

                }#=>if $csvFile.name NU*
                elseif ($csvFile.name -like "CU*") {
                    Write-Debug "This is a 'change user' request so we are making these changes immediately. We will NOT schedule these types of requests and will ignore CSV 'startdate' field."
                    $changeUserAD = @{
                        'SamAccountName'            = $SAM
                        'UserPrincipalName'         = $UPN
                        'Company'                   = $Company
                        'EmailAddress'              = $Email
                        'GivenName'                 = $FirstName
                        'Surname'                   = $LastName
                        'AccountExpirationDate'     = $oEndDate
                    }#=>$changeUserAD
                    try {
                        $oChangeADUser = Get-ADUser -Filter {mail -eq $Email} -ErrorAction 'Stop' -WarningAction 'Stop'
                        Set-ADUser $oChangeADUser @changeUserAD -ErrorAction 'Stop' -WarningAction 'Stop'
                    }
                    catch {
                        Write-Debug "Unable to change user $($FullName) in AD. Error is `n`n $($_.Exception.Message)"
                        Write-CustomEventLog -message "Unable to modify AD User $($Fullname) with SAM $($SAM) in AD.  Full error details below; `n`n $($_.Exception.Message)" -entryType "Warning"
                    }#=>try/catch $oChangeADUser

                    if(-not($oChangeADUser)) {
                        Write-Debug "Unable to change user $($FullName) in AD."
                        Write-CustomEventLog -message "Unable to modify AD User $($Fullname) with SAM $($SAM) in AD.  Full error details below; `n`n $($oChangeUser.Error)" -entryType "Warning"
                    } else {
                        #change user request was fine...
                        Write-Debug "Successfully changed AD user $($FullName)"
                        #$successUsers += -join($FullName,",",$SAM,",","Successfully changed AD user $($FullName)")
                        Write-CustomEventLog -message "Successfully updated AD User Name: $($FullName) - SAM: $($SAM) in AD. Details are below. `n`n $($changeUserAD | Out-String)" -entryType "Information"
                    }#=> if not $oChangeADUser

                }#=> elseif $csvFile.name -like CU*
            }#=>ForEach $user !$isScheduled

            Write-Debug "Renaming our current csv file $($csvFile.FullName) and addding a .done extension. Also making the file read-only."
            Send-CustomMail -setMailType "WARNING"
            Rename-Item -Path $csvFile.FullName -NewName "$($csvFile.FullName).done" -Force
            Set-ItemProperty -Path "$($csvFile.FullName).done" -name IsReadOnly -Value $true

        }#=>foreach $csvFile
    }#=>if $csvFiles
    else {
        Write-Debug "No CSV files found in $($csvPath) that require processing.  Nothing to do this round."
        Stop-Transcript
        exit
    }#=>else $csvFiles
}#=>if !$isScheduled
else {
    Write-Debug "This is a scheduled task to create a new user.  Let's build our request and create the user."
    #Checking to see if a user already exists in AD with the same email address...
    try {
        Get-ADUser -Filter {mail -eq $pEmail}
    }
    catch {
        
    }
    if (Get-AdUser -Filter {mail -eq $pEmail}) {
        Write-Debug "A user with email address $($pEmail) already exists in AD.  We cannot add $($pFullName) with $($pEmail)."
        Write-CustomEventLog -message "A user with email address $($pEmail) already exists in AD.  Skipping the creation of user $($pFullName) with SAM $($pSAM)" -entryType "Warning"
    }#=if get-aduser
    else {
        Write-Debug "No existing user in AD with email address $($pEmail) so we can create our user."
        Write-Debug "Attempting to get properties of our template user $($pCopyUser) to copy from..."

        try {
            $templateUser = Get-ADUser -filter {name -eq $pCopyUser} -Properties MemberOf,EmailAddress -ErrorAction 'Stop' -WarningAction 'Stop'    
        }
        catch {
            Write-Debug "We were unable to find the template user $($copyUser) so we have to skip this new AD user and go to the next row in the CSV file."
            Write-CustomEventLog -message "We were unable to find the template user $($copyUser) when attempting to create new user $($FullName) with SAM $($SAM).  Fatal error.  Exiting script." -entryType "Error"
            Stop-Transcript
            exit
        }#=>try/catch $templateUser

        if (-not($templateUser)) {
            Write-Debug "We were unable to find the template user $($pCopyUser) so we cannot create the new user $($pFullName)"
            Write-CustomEventLog -message "We are unable to find the template user $($pCopyUser) in AD.  Unable to create new user $($pFullName) due to this error." -entryType "Error"
            Stop-Transcript
            exit

        } else {
            $Password = (ConvertTo-SecureString -AsPlainText 'Cloudcom.1' -Force)
            $oEndDate = [datetime]::ParseExact(($pEndDate).Trim(), "dd/MM/yyyy", $null) #This conerts to CSV 'EndDate' field from a string to a datetime object which is required for the New-AdUser cmdlet 'AccountExpirationDate' parameter.

            $newUserAD = @{
                'Company'                   = $pCompany
                'AccountExpirationDate'     = $oEndDate
                'ChangePasswordAtLogon'     = $true
                'Enabled'                   = $true
                'PasswordNeverExpires'      = $false
            }#=>$newUserAD

            $newUserExch = @{
                'SamAccountName'            = $pSAM
                'UserPrincipalName'         = $pUPN
                'Name'                      = $pFullName
                'PrimarySmtpAddress'        = $pEmail
                'FirstName'                 = $pFirstName
                'LastName'                  = $pLastName
                'Password'                  = $Password
            }

            #Let's get the OU that our template user belongs to and apply that to our new user...
            $OU = ($templateUser.DistinguishedName).Substring(($templateUser.DistinguishedName).IndexOf(",")+1)
            Write-Debug "Our OU for new user $($pFullName) is $($OU) from copy of our template user $($pCopyUser) with OU of $($templateUser.DistinguishedName)"

            #Let's update our $newUserExch properties with this OU...
            $newUserExch['OrganizationalUnit'] = $OU


            try {
                Add-PSSnapin "*Exchange*"
            }
            catch {
                Write-Debug "Unable to connect to Exchange PowerShell due to the following error $($_.Exception.Message).  This is likely a fatal error for the entire email portion of the script."
                Write-CustomEventLog -message "Unable to connect to Exchange Powershell to create mailbox for user $($FullName) due to the following error: `n $($_.Exception.Message) `n This is likely a fatal error for the entire email portion of the script.  This error should be remedied or no email boxes will be created." -entryType "Error"
                exit
            }
            try {
                Write-Debug "Running Get-Mailbox using identity parameter of $($templateUser.EmailAddress)"
                $copyMailProps = Get-MailBox -Identity $($templateUser.EmailAddress) -ErrorAction 'Stop' -WarningAction 'Stop' | Select-Object AddressBookPolicy,Database
            }
            catch {
                Write-Debug "Unable to Get-Mailbox for template user $($templateUser.EmailAddress) which means we are unable to activate $($pFullName) in AD or Exchange. Continuing to next user."
                Write-CustomEventLog -message "Unable to Get-Mailbox for template user $($templateUser.EmailAddress) which means we are unable to activate $($pFullName) in AD or Exchange." -entryType "Warning"
                Stop-Transcript
                exit
            }#=>try/catch CopyMailProps

            if (-not($copyMailProps)) {
                Write-Debug "Unable to Get-Mailbox for template user $($templateUser.EmailAddress) which means we are unable to activate $($pFullName)'s Exchange Email. Continuing to next user."
                Write-CustomEventLog -message "Unable to Get-Mailbox for template user $($templateUser.EmailAddress) which means we are unable to activate $($pFullName)'s Exchange Email." -entryType "Warning"
                exit
            } else {
                #We got our $copyMailProps properties for Database and AddressBookPolicy so we'll just add that to the $newUserExch hashtable.
                Write-Debug "`$copyMailProps has returned; $($copyMailProps | Out-String)"
                $newUserExch['AddressBookPolicy'] = $copyMailProps.AddressBookPolicy
                $newUserExch['Database'] = $copyMailProps.Database
            }#=>if not $copyMailProps

            try {
                $oNewExchUser = New-Mailbox @newUserExch -ErrorAction 'Stop' -WarningAction 'Stop'
            }
            catch {
                Write-Debug "Unable to create new user $($pFullName) using New-Mailbox.  Error message `n`n $($_.Exception.Message)"
                Stop-Transcript
                exit
            }
            if(-not($oNewExchUser)) {
                Write-Debug "Something went wrong with adding our new $($pFullName) user to AD and Exchange. `n`n $($_.Exception.Message)"
                Write-CustomEventLog -message "We were unable to add our new user $($pFullName) to AD and Exchange. Full error details below; `n`n $($_.Exception.Message)." -entryType "Warning"
                Stop-Transcript
                exit
            }
            #Adding user went well now let's update the AD properties for this user that can't be done using the New-Mailbox cmdlet.
            Write-Debug "We created our new user $($pFullName) in AD and Exchange. Modifying AD user properties."
            try {
                $setUserADProps = Set-ADUser @newUserAD -ErrorAction 'Stop' -WarningAction 'Stop'
            }
            catch {
                Write-Debug "Unable to modify AD user properties for $($pFullName).  Continuing to next user."
                Write-CustomEventLog -message "We were unable to modify AD properties for user $($pFullName).  Full error is `n`n $($_.Exception.Message).`n`n User properties we want to modify are $($newUserAD | Out-String)" -entryType "Error"
                Stop-Transcript
                exit
            }#=> try/catch $setUserADProps
            if(-not($setUserADProps)) {
                Write-Debug "Unable to modify AD user properties for $($pFullName).  Continuing to next user."
                Write-CustomEventLog -message "We were unable to modify AD properties for user $($pFullName).  Full error is `n`n $($_.Exception.Message).`n`n User properties we want to modify are $($newUserAD | Out-String)" -entryType "Error"
                Stop-Transcript
                exit
            } else {
                Write-Debug "Successfully created new Exchange mailbox and modified AD properties for user"
                Write-CustomEventLog -message "Successfully created new AD User and Exchange Mailbox for $($pFullName).  AD and Exchange Details included below; `n`n $($newUserExch | Out-String) `n`n $($newUserAD | Out-String)" -entryType "Information"
            }
        }#=>if/else $templateuser
    }#=>else get-ADUser
}#=>if isScheduled
Stop-Transcript
$transcriptContent = Get-Content -Path $TranscriptLog -RAW
Write-CustomEventLog -message "Finished running script. Full transaction log details are below; `n`n` $($transcriptContent)" -entryType "Information"