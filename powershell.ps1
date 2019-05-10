$DebugPreference = "Continue"
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
   Write-Debug "First Name (CSV): $($user.Firstname)"
   Write-Debug "Last Name (CSV): $($user.Lastname)"
   Write-Debug "Company (CSV): $($user.Company)"
   Write-Debug "Email/Domain (CSV): $($user.'Email/Domain')"
   Write-Debug "EndDate (CSV): $($user.EndDate)"
   #etc.etc.

  $FirstName = Format-CsvValue -isTitleCase $true -sValue $User.FirstName #using our new function let's properly format the firstname.
  Write-Debug "First Name (Script): $FirstName"

  $LastName = Format-CsvValue -isTitleCase $true -sValue $User.LastName
  Write-Debug "Last Name (Script): $LastName"

  $FullName = -join($($User.FirstName)," ",$($User.LastName)) #not using Format-CsvValue here because that was already done for $FirstName and $LastName vars seperately
  Write-Debug "Full Name (Script): $FullName"

  $SAM = $(-join (($user.FirstName).Substring(0,1),$user.LastName).ToLower())
  $SAM = $SAM.ToLower()
  Write-Debug "SAM (Script): $SAM"


  
  $dnsroot = "@$((Get-ADDomain).dnsroot)"
  Write-Debug "DNS Root (Script): $dnsroot"

  $Password = (ConvertTo-SecureString -AsPlainText 'Cloudcom.1' -Force)

  $UPN = -join ($SAM,$dnsroot)
  Write-Debug "UPN (Script): $UPN"

  $OU = $user.OU #uncommented and changed to read the "OU" header in the CSV file.
  Write-Debug "OU (Script): $OU"

  $email = -join ($Sam,$dnsroot)
  Write-Debug "Email (Script): $Email"
  #your CSV file contains headers that you aren't bringing into the script
  #Examples;
  #missing 'Company'
  #missing StartDate
  #missing EndDate

Try {
    if (!(get-aduser -Filter {samaccountname -eq "$SAM"})){
      #your CSV file contains headers that do not exist in your $Parameters hash table and they are as follows;
      #no 'Company' mapping.
      #no 'StartDate' mapping.
      #no 'EndDate' mapping.
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

    $oNewUser = New-ADUser @Parameters

    Write-Verbose "[PASS] Created $FullName"
    $successUsers += "$FullName , $SAM"
} Catch {
  $ErrorMessage = $_.Exception.Message
  $FailedItem = $_.Exception.ItemName
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
Invoke-Item $LogFolder #is this to open the Widnows Explorer view of your log folder? yes, for both the failed and succesfully created users
