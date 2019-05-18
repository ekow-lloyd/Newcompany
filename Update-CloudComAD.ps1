<#
.SYNOPSIS
Create a new or modify an existing CloudCom user in the CloudCom AD.

.DESCRIPTION
Create a new or modify an existing CloudCom user in the CloudCom AD.
Two parameter sets ("Init" and "Scheduled") exist.  This is to make it easier to call the script when it's initially called (when reading from a CSV) or when it's called as part of a scheduled task.
Only paramaters that are members of a parameter set can be called in an single instance.

This is used so the script can read the CSV file (init) and process the request based on the startdate value in the read CSV file.
If the startdate of the user is within 48 hours of the scrpit run then it'll automatically add the user to AD at the time of script run.
Otherwise, if the startdate of the user is beyond 48 hours of the script run, the script will create a scheduled tasks to add the user within 48 hours of the start date.

.PARAMETER isScheduled
Type: Boolean ($true or $false)
Mandatory: Yes
Set: Init, Scheduled
Tells the script whether or not to run in a scheduled task mode ($true) or 'input from csv' mode ($false)


.PARAMETER RequestType
Type: String
Mandatory: Yes
Set: Init, Scheduled
Validation: "New", "Change"

Tells the script if the request type is a new user (-RequestType New) or a change user (-RequestType Change).
This parameter is mandatory and only accepts "New" or "Change" values.  Any other values passed to this paramater will cause the script to not run.

.PARAMETER sFirstName
Type: String
Mandatory: Yes
Set: Scheduled

The firstname of the user.  Supplied as a parameter and value to the script when run from a scheduled task.

.PARAMETER sLastName
Type: String
Mandatory: Yes
Set: Scheduled

The last name of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER sSAM
Type: String
Mandatory: Yes
Set: Scheduled

The SamAccountName of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER sUserName
Type: String
Mandatory: Yes
Set: Scheduled

The username of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER sOU
Type: String
Mandatory: Yes
Set: Scheduled

The OU the user will belong to.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER sStartDate
Type: String
Mandatory: Yes
Set: Scheduled

The Start Date of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER sEndDate
Type: String
Mandatory: Yes
Set: Scheduled

The End Date of the user.  Supplied as a parameter and value to the scirpt when run from a scheduled task.

.PARAMETER sCompany
Type: String
Mandatory: Yes
Set: Scheduled

The Company the user belongs to. Supplied as a parameter and value to the scirpt when run from a scheduled task.

.INPUTS
When run in "Init" set the path of the CSV file(s) are required.

.OUTPUTS
Outputs two log files to the running user's Desktop ($env:username\desktop\)

.EXAMPLE
Run the script to pull the CSV file only.

PS> New-CloudComUser.ps1 -requestType New -isScheduled $false

.EXAMPLE
Run the script as a future dated scheduled task.

PS> New-CloudComUser.ps1 -requestType New -isScheduled $true -sFristName "John" -sLastName "Doe" -sSAM "jdoe" -sUserName "john.doe" -sOU "OU=USERS,DC=AD,DC=local" -sStartDate "12/05/2019" -sEndDate "12/05/2020" -sCompany "ABC Co."

#>
Param(
    # Whether or not this is a scheduled task...
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")] #include $isScheduled in both (scheudled and init) parameter sets.
    [Parameter(Mandatory=$true,ParameterSetName="Init")]
    [bool]
    $isScheduled,
    # Whether or not this is a "new user" request or "change user" request.
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")] #include $requestType in both (scheudled and init) parameter sets.
    [Parameter(Mandatory=$true,ParameterSetName="Init")]
    [ValidateSet('new','change')] #only two options can be passed to this parameter "new" or "changed" and it is required.
    [string]
    $requestType,
    # Let's define all required parameters when creating a user when it's a scheduled task.  Scheduled tasks require additional parameters because the initial CSV that was loaded will no longer be used.  Instead, all values from the CSV will be stored as arguments (parameters) to the script.
    # First Name
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $sFirstName,
    # Last Name
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $sLastName,
    # SAM
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $sSAM,
    # Username
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $sUsername,
    # OU
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $sOU,
    # start date
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $sStartDate,
    # end date
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $sEndDate,
    # company
    [Parameter(Mandatory=$true,ParameterSetName="Scheduled")]
    [string]
    $sCompany
)

$DebugPreference = "Continue"
$VerbosePreference = "Continue"
$ErrorActionPreference = "Stop"
$LogFolder = "$env:userprofile\desktop\logs" #log file location.


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

if (!(Get-module ActiveDirectory )) { #checking if the ActiveDirectory module is loaded...
    Write-Debug "ActiveDirectory module not loaded. Importing..."
    Import-Module ActiveDirectory
    Clear-Host
} #=if get-Module

if (!($isScheduled)) {
    Write-Debug "This is not a scheduled task so we can safely assume this is an initial read of a CSV file.  Looking for which CSV file to process..."

    switch ($requestType) {
        "New" { $csvPath = "C:\temp\new_cloudcom.csv" } #changeme
        "Change" {$csvPath = "C:\temp\change_cloudcom.csv"} #changeme
        Default {
            #This is a problem.  $requestType is a mandatory field and it should have a value of New or Change but it doesn't so this is a hard error. Write the error and quit the script.
            Write-Output "Error determining requestType which means we don't know which CSV file to choose.  This is a fatal error." | Out-File "$LogFolder\error.log"
            Throw "Error determining requestType which means we don't know which CSV file to choose.  This is a fatal error."
        }
    }

    #import our CSV file.
    try {
        $Users = Import-CSV $csvPath
    }
    catch {
        Write-Output "Unable to import our CSV file: $csvPath. This is a fatal error with error: $Error[0].Exception.Message" | Out-File "$LogFolder\error.log"
        Throw $Error[0].Exception.Message
    }#=> try $Users

    #imported our CSV file properly.  Let's process the file for new users...
    ForEach ($User in $Users){
        #debugging purposes...
        Write-Debug "First Name (CSV): $($User.Firstname)"
        Write-Debug "Last Name (CSV): $($User.Lastname)"
        Write-Debug "SAM (CSV): $($User.Sam)"
        Write-Debug "OU (CSV): $($User.OU)"
        Write-Debug "StartDate (CSV): $($User.startdate)"

        Write-Debug "End Date (CSV): $($User.enddate)"
        Write-Debug "Company (CSV): $($User.Company)"
        #=>debugging purposes.

        #Let's properly format all the values in this *ROW* of the CSV. Trim() where necessary and change to Title Case where necessary - also create a new variable so we can use it later when creating the user in AD using the New-ADuser cmdlet.
        $FirstName = Format-CsvValue -isTitleCase $true -sValue $User.FirstName #trim and title case
        $LastName = Format-CsvValue -isTitleCase $true -sValue $User.LastName #trim and title case.
        $SAM = Format-CsvValue -isTitleCase $false -sValue $User.Sam #trim only.
        $Username = (Format-CsvValue -isTitleCase $false -sValue $user.Username).ToLower #trim and make lowercase.
        $OU = (($user.OU).Trim()).Replace('"','') # Trim whitespace and delete the double quotes that may exist in the CSV file.
        $StartDate = Format-CsvValue -isTitleCase $false -sValue $User.startdate #trim only.
        $EndDate = Format-CsvValue -isTitleCase $false -sValue $User.enddate #trim only.
        $Company = Format-CsvValue -isTitleCase $false -sValue $User.company #trim only since company names are rather specific on how they're spelled out.
        #=> End of CSV values.

        #Let's build other necessary variables that are required parameters for the New-ADuser cmdlet out of the information provided by the CSV file or other sources...
        $FullName = -join($($FirstName)," ",$($LastName)) #join $Firstname and $Lastname and a space to get FullName
        $DNSroot = "@$((Get-ADDomain).dnsroot)"
        $UPN = -join($Username, $dnsroot)
        

        $oStartDate = [datetime]::ParseExact(($User.StartDate).Trim(), "dd/MM/yyyy", $null) #This converts the CSV "startdate" field from a string to a datetime object.
        Write-Debug "StartDate Object (Script): $oStartDate"


        

        $FullName = -join ($($User.FirstName), " ", $($User.LastName)) #not using Format-CsvValue here because that was already done for $FirstName and $LastName vars seperately




    }#=>ForEach $user !$isScheduled

}#=>if !$isScheduled

$a = 1;
$b = 1;
$failedUsers = @()
$successUsers = @()
