<#Mike in IT Comments 4/25/2019 @ 4:15PM ET
* Still need to figure out how to differentiate between a "changeuser" and a "newuser" request and script for that.  Are you going to do it by filename?
  The script as-is will treat all CSV imports as a new user.

* Think about 'collision rules' when modifying and/or creating new users.  If you already have a John Doe user working for XYZCorp in AD and then a new...
  user with that same name of John Doe starts at XYZCorp your script will fail because previous user is taken.
  You can either put logic in your script to detect this and put in 'collision rules' (e.g. if John Doe exists 3 times and #4 is hired add a 4 to the end of their name)...
  Another option is you can just rely on your output file notifying you of this issue and you can manually add the user using standard tools (Active Diretory Users and Computers snap-in).

* I don't think your try/catch will work as anticipated - have you tested this? Have you 'forced' an error to make sure it is catching New-ADUser issues?

* One thing I never thought to ask you - are all of your domain controllers running Windows Server 2012 or better?

* The CSV provided today has "NewUser" but no CopUserTemplate, DB, or AddressBookPolicy defined fields.  Why wouldn't you want to use the ''Copy user template' on NewUser requests?

#>

function Format-CsvValue {
  [CmdletBinding()]
  param (
    #Sets whether or not we want to format the provided string into 'title' (aka Proper) case when using named values.
    #When isTitleCase = $true the function will take the input string ($sValue) and format it to proper(title) case and will also remove leading and trailing whitespaces.  Example; "JoHN SmITH" will return "John Smith" or "   JaNE " will return "Jane" (removed whitespaces and set to title case).
    [Parameter(Mandatory=$false)]
    [bool]
    $isTitleCase = $false,
    #The string value that's passed into the function to properly format.
    #Example: Format-CsvValue -isTitleCase $true -sValue $mvar
    #Example: To only remove whitespace from a string-> Format-CsvValue -sValue $myvar
    [Parameter(Mandatory=$true)]
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
    } else {
      #only whitespace trim is required, process that request.
      $rValue = $sValue.Trim() #Remove leading/trailing whitespace.
    }#=>if/isTitleCase
  }#=>process
  
  end {
    #return the value through the function.
    $rValue
  }
} #=>Format-CsvValue

If (!(Get-module ActiveDirectory )) {
    Import-Module ActiveDirectory
    Clear-Host
} #=if get-Module
  
$Users=Import-csv c:\cloudcom2.csv
$a=1;
$b=1;
$failedUsers = @()
$successUsers = @()
$VerbosePreference = "Continue"
$ErrorActionPreference='stop'
$LogFolder = "$env:userprofile\desktop\logs"

ForEach($User in $Users) {
  
  #$FirstName = $User.FirstName.substring(0,1).toupper()+$User.FirstName.substring(1).tolower()
  $FirstName = Format-CsvValue -isTitleCase $true -sValue $User.FirstName #using our new function let's properly format the firstname.
  #$LastName  = $User.LastName.substring(0,1).toupper()+$User.LastName.substring(1).tolower()
  $LastName = Format-CsvValue -isTitleCase $true -sValue $User.LastName

  $FullName = $User.FirstName + " " + $User.LastName #not using Format-CsvValue here because that was already done for $FirstName and $LastName vars seperately

  $SAM = ($user.FirstName.Substring(0,1) + $user.LastName).ToLower #no need to put this on a new line itself do it all in a single step.

  #$SAM=$sam.tolower() #I moved this variable up a couple of lines. I like to keep variables together especially when doing multiple operations on the same variable so then you know what happened later :)
  
  $dnsroot = '@' + (Get-ADDomain).dnsroot

  $Password = (ConvertTo-SecureString -AsPlainText 'Cloudcom.1' -Force)

  $UPN = $SAM + "$dnsroot" #why are you quoting "$dnsroot" but not $SAM? Why are you quoting at all? :)

  $OU = $user.OU #uncommented and changed to read the "OU" header in the CSV file.

  $email = $Sam + "$dnsroot" #same comment as for $UPN
  #your CSV file contains headers that you aren't bringing into the script
  #Examples;
  #missing 'Company'
  #missing StartDate
  #missing EndDate

Try {
    if (!(get-aduser -Filter {samaccountname -eq "$SAM"})){
      $Parameters = @{
        'SamAccountName'        = $Sam
        'UserPrincipalName'     = $UPN 
        'Name'                  = $Fullname
        'EmailAddress'          = $Email 
        'GivenName'             = $FirstName 
        'Surname'               = $Lastname  
        'AccountPassword'       = $password 
        'ChangePasswordAtLogon' = $true 
        'Enabled'               = $true 
        'Path'                  = $OU
        'PasswordNeverExpires'  = $False
        #your CSV file contains headers that do not exist in your $Parameters hash table and they are as follows;
        #no 'Company' mapping.
        #no 'StartDate' mapping.
        #no 'EndDate' mapping.
      } #=>Parameters

      <#
      #This section may no longer be necessary if you are going to be copying a user 'template' from AD vs settign it in code...
      #If you still want to maintain the the 'template' in code vs in AD then we'll need to have a "Type" field defined in the CSV file.
      switch ($user.Type) { 
          "Staff" { 
            $Parameters['Path'] = "OU=Staff,OU=Cloudcom,DC=E-L,DC=local"
            }
          "Admin" { 
            #Same as above with 'staff' but for Administration users.
            $Parameters['Path'] = "OU=Administration,OU=Cloudcom,DC=E-L,DC=local"
          }
      #>
    } #=>if get-aduser

    New-ADUser @Parameters
    Write-Verbose "[PASS] Created $FullName "
    $successUsers += $FullName + "," +$SAM
} Catch {
  Write-Warning "[ERROR]Can't create user [$($FullName)] : $_"
  $failedUsers += $FullName + "," +$SAM + "," +$_
}#=>catch
} #=>ForEach User
if ( !(test-path $LogFolder)) {
    Write-Verbose "Folder [$($LogFolder)] does not exist, creating"
    new-item $LogFolder -type directory -Force 
} #=> test-path LogFolder


Write-verbose "Writing logs"
$failedUsers  | ForEach-Object {"$($b).) $($_)"; $b++} | out-file -FilePath  $LogFolder\FailedUsers.log -Force -Verbose
$successUsers | ForEach-Object {"$($a).) $($_)"; $a++} | out-file -FilePath  $LogFolder\successUsers.log -Force -Verbose

$su=(Get-Content "$LogFolder\successUsers.log").count
$fu=(Get-Content "$LogFolder\FailedUsers.log").count


Write-Host "$fu user creation unsuccessful " -NoNewline -ForegroundColor red
Write-Host "$su Users Successfully Created "  -NoNewline -ForegroundColor green
Write-Host " Review LogsFolder" -ForegroundColor Magenta
Start-Sleep -Seconds 5
Invoke-Item $LogFolder #is this to open the Widnows Explorer view of your log folder?