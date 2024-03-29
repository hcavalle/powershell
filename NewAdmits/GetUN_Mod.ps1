Set-PSDebug -strict

        [int]      $Year = 2016
        [string]   $connStr = "data source=UCLAWSQL1;initial catalog=Admissions;Integrated Security=true;"
        [string]   $cmdStr = "BulkAddAdmits" 
        [string]   $admittedDn = "OU=Admitted Students,OU=Network Users,DC=lawnet,DC=lcl"
        [string]   $studentDn = "OU=Students,OU=Network Users,DC=lawnet,DC=lcl"
        [string]   $homeRoot = "\\uclawdata1\StuHome\"
        [string]   $googleAppDn = "CN=Google Apps Users,OU=Network Users,DC=lawnet,DC=lcl"
        [string]   $admittedNotComingOu = "OU=Admitted Students - Not Coming,OU=Network Users,DC=lawnet,DC=lcl"
        [DateTime] $moveStart = "8/1/" + [DateTime]::Now.Year.ToString()
        [DateTime] $moveEnd = "9/1/" + [DateTime]::Now.Year.ToString()
        [string]   $upnDomain = "lawnet.lcl"
        [string]   $emailDomain = "lawnet.ucla.edu"

###Copy functions from LawAccount.psm1 start with check length, ending with getUserName over to here ######################

######Harrison ADDED FUNCTIONS #######################################

function checkSpecChar
{
    param(
        [string]$Name
    )
    if ($Name.IndexOf(" ") -gt 0){
           [int]$posSpace = $Name.IndexOf(" ")
           #$LastName.Remove($pos)
           $leftPart = $Name.Substring(0, $posSpace)
           #$rightPart = $LastName.Substring($posSpace+1)
           $Name = "$leftPart"
           #DEBUG Write-Host "$posSpace $leftPart $rigthPart New lastname: $Name"
        }
     if ($Name.IndexOf("-") -gt 0){
           [int]$posDash = $Name.IndexOf("-")
           #$LastName.Remove($pos)
           $leftPart = $Name.Substring(0, $posDash)
           #$rightPart = $LastName.Substring($posDash+1)
           $Name = "$leftPart"
           #DEBUG Write-Host "$posDash $leftPart $rigthPart New lastname: $Name"
        }#>
        
    return $Name
}

function CheckLength{
    #checks the number of repititions of the FI. If more than 1x, try next naming convention. 
    #returns number of FI reps
    
    param (
        [string]$LastName = $(throw "LastName is required."),
        [string]$FirstName = $(throw "FirstName is required."), 
        [string]$CurrentUserName = $(throw "CurrentUserName is required.")
      )
      
      #Check to ensure FN and LN are not whitespace
    if ($FirstName -eq ' ' -or $LastName -eq ' ' -or $Year -lt 1900){
        Write-Host "`rBlank Name or UN, username will stay: " $CurrentUserName
        return 0
    }#>
    while($LastName.contains(" ") -or $LastName.contains("-")){
        #try {$LastName = checkSpecChar($LastName)}
        #catch{ return $CurrentUserName.ToUpper()}
        try {$LastName = checkSpecChar($LastName)}
        catch {return 0}
        #Write-Host "Last Name is: $LastName"
    }#>
      
    try {
        #[string]$FiReps = $CurrentUserName.Substring($LastName.Length)
        $NameLength = $LastName.Length
        [string]$FiReps = $CurrentUserName
        $FiReps = $FiReps.Substring($NameLength)
        #DEBUG Write-Host "FIreps are: $FIReps"
        
    }
    catch{
      return 0
    }
       
    $i = 0
    if ($FiReps.length -lt 6){
      return $i
    }
      
    elseif ($FiReps[0] -eq $FiReps[1]){
      $i = 1
      while ($FiReps[0] -eq $FiReps[$i+1]){
          $i++
          Write-Host $Fireps[$i] 
      }
        Write-Host "FI Reps for $CurrentUserName are " $i 
     }
      
     if($i -gt 0){
        
        Write-Host "FI Reps on $FIReps are " $i
     }
     return $i 
}

function CheckUnique{
    #returns $true if Username is unique, $false if there is a conflict  
    param(
        [string] $TestUserName = $(throw "UserName is required")
    )
    
    #Use system.directorysearcher object to do LDAP query
    $strFilter = "(sAMAccountName=$TestUserName)"
    $objDomain = New-Object System.DirectoryServices.DirectoryEntry #("LDAP://dc=lawnet, dc=lcl")
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.Filter = $strFilter
    $objSearcher.SearchScope = "Subtree"
    #$Result = $objSearcher = New-Object System.DirectoryServices.SearchResult
    #$objSearcher.FindAll()
    
    if ($objSearcher.FindOne() -ne $null){
        #DEBUG Write-Host "Found one" 
        return $false    
    }
    
    if ($TestUserName.Length -gt 15){
        return $false
    }
    
    else{
        return $true
        #DEBUG Write-Host "Is Unique" #DEBUG
    }
    
}

function Email{

    param(
        [string]$emailto,
        [string]$subject,
        [string]$body
    )

    $emailFrom = "ishelp@law.ucla.edu"
    #$emailTo = "help@law.ucla.edu"
    #$subject = "New LawNET account: " + $UserName
    #$body = "Account has excessive repetition of first initial. Change the account ASAP before they access it. You can tell if they have by seeing if it is expired or not. If it is not expired, reach out to them and let them know we can change it. Username: "+$UserName
    $smtpServer = "smtp.ucla.edu"
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($emailFrom, $emailTo, $subject, $body)

}

function GetYear{
    #iterate through until int, get index, substring to get year
    param(
        $CurrentUserName
        )
        
    $j=0
    #$obj = New-Object "System.Double"

    While($j -lt $CurrentUserName.Length){
        if ($CurrentUserName[$j] -eq '2'-or $CurrentUserName[$j] -eq '1'){
            break
        }
        $j++   
    }
    if (($Year = $CurrentUserName.Substring($j)) -gt 1900){
        return $Year
    }
    return 0
    
}

function GetUserName{
    param(
        [string] $CurrentUserName = $(throw "username is required"),
        [string] $LastName = $(throw "last name is required"),
        [string] $FirstName = $(throw "first name is required"),
        [string] $MiddleName,
        [int] $Year = $(throw "year is required")
    )
    
    [string] $UserName = $CurrentUserName
    $FIcount = CheckLength $LastName $FirstName $CurrentUserName
    
    <#Check if FI is repeated more than 1x, if not check if unique and return or move on
    if ( !$FIcount -or $FIcount -lt 2){
           return $UserName.ToUpper()
        
    }#>
    
    #Try lastnameFIMIYear
    if ($MiddleName[0] -ne " " ){
        $UserName = $LastName + $FirstName[0] + $MiddleName[0] + $Year
        if (CheckUnique($UserName)){
            return $UserName.ToUpper()
        }
    }#>
    
    #Try last.firstYear
    if ((CheckUnique($UserName) -eq $false)){
        $UserName = $LastName + "." + $FirstName + $Year
        return $UserName.ToUpper()
    }
    
    #Try lastfirstYear
    elseif ((CheckUnique($UserName) -eq $false)){
    $UserName = $LastName + $FirstName + $Year
        return $UserName.ToUpper()
    }
    
    #Try lasname+each letter into first+year (ex: smithjo2013 then smithjoh2013)
    else {
        try{
            $PartFirst = $FirstName[0]
            for ($i=1; $i -lt $FirstName.Length; $i++){
                $PartFirst += $FirstName[$i]
                $UserName = $LastName + $PartFirst + $Year
                if (CheckUnique($UserName)){
                    return $UserName.ToUpper()
                }
            }
            
        }
       catch{
            "Error parsing length on first name"
            $error[0]
       }#>
   
    }
       
    #EMAIL HELP IF STILL REPETITION OF FI IS GT 3
    if ($FIcount -gt 3){
        $subject = "New LawNET account: " + $UserName
        $body = "Account has excessive repetition of first initial. Change the account ASAP before they access it. You can tell if they have by seeing if it is expired or not. If it is not expired, reach out to them and let them know we can change it. Username: "+$UserName
        Email "help@law.ucla.edu" $subject $body
        #DEBUG Write-Host "Emailed Help"
    }
    #returns original username since we know this is unique
    #DEBUG Write-Host "returning default"
    return $CurrentUserName.ToUpper()
}

###################end of harrison ADDED FUNCTIONS ############################>

#### TESTING GETUSERNAME FUNCTION########

#Import-Module -Name "C:\ScheduledTasks\NewAdmits\LawAccount.psm1"

# test all students
$AlumOU = [ADSI]"LDAP://OU=Alumni,OU=Network Users,DC=lawnet,DC=lcl"

foreach($child in $AlumOU.Children)
{
        [string]$CurrentUserName = $child.sAMAccountName
        
        $FirstName = $child.givenName
        $LastName = $child.sn
        $MiddleName = $child.middleName
        #DEBUG Write-Host "Calling Get year"
        [int]$Year = getYear($CurrentUserName)
        # DEBUG Write-Host $Year
        if($child.sAMAccountName -eq "abdollahi2009"){
            
            #[string]$CurrentUserName = "$LastName$Year"
        
            Write-Host "OG Username" $CurrentUserName
            [string] $NewUserName = GetUserName $CurrentUserName $LastName $FirstName $MiddleName $Year

            Write-Host "Refined UN:" $NewUserName 
            #Write-Host $FirstName 
            #Write-Host $LastName 
            #Write-Host $MiddleName 
            #Write-Host $Yearwinw
        }
}
#>

Write-Host $error[0]
$error.clear

<#testcase1
$CurrentUserName = "wu2013"
$FirstName = "Andrew"
$LastName = "Wu"
$MiddleName
#Write-Host "Calling Get year"
[int]$Year = getYear($CurrentUserName)
#Write-Host $Year 
Write-Host "OG Username" $CurrentUserName
$CurrentUserName = GetUserName $CurrentUserName $LastName $FirstName $MiddleName $Year
Write-Host "Refined UN:" $CurrentUserName 
#$error[0]
#Write-Host $FirstName 
#Write-Host $LastName 
#Write-Host $MiddleName 
#Write-Host $Year
#>

<# Testing
1. LastF2013
2. LastFM2013
3. LAST.FIRST2013
4. LASTFIRST2013
5. LASTFI...FIR...FIRS...2013
6. Return OG UN
#>
