function Add-UserGroup
{
    param(
        [string] $userDn = $(throw "user DN is required"),
        [string] $groupDn = $(throw "group DN is required")
    )
    
    $group = [ADSI]("LDAP://$groupDn")
        
    if(!($group.IsMember("LDAP://$userDn")))
    {
        Write-Host "`r`tAdd to group ($($group.name))"
        $group.Add("LDAP://$userDn")
    }
}

function Check-UserGroup 
{
	param(
        [string] $userDn = $(throw "user DN is required"),
        [string] $groupDn = $(throw "group DN is required")
    )
	
	$group = [ADSI]("LDAP://$groupDn")
        
    return $group.IsMember("LDAP://$userDn")
}

function Remove-UserGroup
{
    param(
        [string] $userDn = $(throw "user DN is required"),
        [string] $groupDn = $(throw "group DN is required")
    )
    
    $group = [ADSI]("LDAP://$groupDn")
        
    if($group.IsMember("LDAP://$userDn"))
    {
        Write-Host "`r`tRemove from group ($($group.name))"
        $group.Remove("LDAP://$userDn")
    }
}

######Harrison ADDED FUNCTIONS #######################################

function CheckLength{
    #checks the number of repititions of the FI. If more than 1x, try next naming convention. 
    #returns number of FI reps
    
    param (
        [string]$LastName = $(throw "LastName is required."),
        [string]$FirstName = $(throw "FirstName is required."), 
        [string]$CurrentUserName = $(throw "CurrentUserName is required.")
      )
      
          #Check to ensure FN and LN are not whitespace
    if ($FirstName -eq ' ' -or $LastName -eq ' '){
        $ErrMessage = "`rBlank Name or UN, or year is off. Username will stay: "+$CurrentUserName
        Write-Host $ErrMessage 
        return 0
    }#>
      
      try {
        #[string]$FiReps = $CurrentUserName.Substring($LastName.Length)
        $NameLength = $LastName.Length
        if($NameLength -gt (($CurrentUserName.Length)-4)){$NameLength = (($CurrentUserName.Length)-4)}
        
        $FiReps = $CurrentUserName
        $FiReps = $FiReps.Substring($NameLength)
        #DEBUG Write-Host "FIreps are: $FIReps"
        
      }
      catch{ 
        $ErrMessage = "`r"+$CurrentUsername +": Trouble checking length of LN. Username will be default."
        Write-Host $ErrMessage
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
        }
      }
      
      if($i -gt 0){
        
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
    $smtpServer = "mail.law.ucla.edu"
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($emailFrom, $emailTo, $subject, $body)

}

function GetYear{
    #iterate through until int, get index, substring to get year
    param(
        $CurrentUserName
        )
        
    $j=0

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

function checkSpecChar
{
    param(
        [string]$Name
    )
    if ($Name.IndexOf(" ") -gt 0){
           [int]$posSpace = $Name.IndexOf(" ")
           $leftPart = $Name.Substring(0, $posSpace)
           #$rightPart = $LastName.Substring($posSpace+1)
           $Name = "$leftPart"
        }
        
     if ($Name.IndexOf("-") -gt 0){
           [int]$posDash = $Name.IndexOf("-")
           $leftPart = $Name.Substring(0, $posDash)
           #$rightPart = $LastName.Substring($posDash+1)
           $Name = "$leftPart"
        }#>
        
    return $Name
}

function GetUserName{
    param(
        [string] $CurrentUserName = $(throw "username is required"),
        [string] $LastName = $(throw "last name is required"),
        [string] $FirstName = $(throw "first name is required"),
        [string] $MiddleName,
        [int] $Year = $(throw "year is required")
    )
    
    if($Year -lt 1900){
        return $CurrentUserName.ToUpper()
    }
    
    while($LastName.contains(" ") -or $LastName.contains("-")){
        #try {$LastName = checkSpecChar($LastName)}
        #catch{ return $CurrentUserName.ToUpper()}
        try {$LastName = checkSpecChar($LastName)}
        catch {return $CurrentUserName.ToUpper()}
    }#>
    
    [string] $UserName = $CurrentUserName
    $FIcount = CheckLength $LastName $FirstName $CurrentUserName
    
    #Check if FI is repeated more than 1x, if not check if unique and return or move on
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
    }#>
    
    #Try lastfirstYear
    elseif ((CheckUnique($UserName) -eq $false)){
    $UserName = $LastName + $FirstName + $Year
        return $UserName.ToUpper()
    }#>
    
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
       }
   
    }
       
    #EMAIL HELP IF STILL REPETITION OF FI IS GT 3
    if ($FIcount -gt 3){
        $subject = "New LawNET account: " + $UserName
        $body = "Account has excessive repetition of first initial. Change the account ASAP before they access it. You can tell if they have by seeing if it is expired or not. If it is not expired, reach out to them and let them know we can change it. Username: "+$UserName
        Email "help@law.ucla.edu" $subject $body
    }
    
    #returns original username since we know this is unique
    return $CurrentUserName.ToUpper()
}

###################end of harrison ADDED FUNCTIONS ############################>

function Disable-Account
{
    param(
        [string] $cn = $(throw "common name is required."), 
        [string] $admittedDn = "OU=Admitted Students,OU=Network Users,DC=lawnet,DC=lcl", 
        [string] $studentDn = "OU=Students,OU=Network Users,DC=lawnet,DC=lcl",
        [string] $googleAppDn = "CN=Google Apps Users,OU=Network Users,DC=lawnet,DC=lcl",
        [string] $admittedNotComingOu = "OU=Admitted Students - Not Coming,OU=Network Users,DC=lawnet,DC=lcl",
        $admitGroups
    )

    Write-Host "`rDisable user ($cn)"

    $disableAccountFlag = 2

    #Write-Host $name
    $AdmitOU = [ADSI]("LDAP://$admittedDn")
    $user = [ADSI]("LDAP://CN=$cn,$admittedDn")
    $userDn = $user.distinguishedName
    
	#If not in Google Groups Delete, check if user is in google group!
	
	if (!(Check-UserGroup $userDn $googleAppDn)){
			Write-Host "Delete not coming accounts"
			$AdmitOU.Invoke("delete", "user", "CN=$cn")
		}
	
	
	#Else
	else{
	
		#check if student has been moved to student OU
		if(!$student.distinguishedName)
		{
			$student = [ADSI]("LDAP://CN=$cn,$studentDn")
			
			if(Test-Path $student.homeDirectory)
			{
				Write-Host "`tDelete user folder $($student.homeDirectory)"
				Remove-Item $student.homeDirectory.ToString() -Recurse
			}
			
			if($student.legacyExchangeDN)
			{
				Write-Host "`tDisable mail user $($student.mail)"
				$dn = $student.distinguishedName.ToString()
				Disable-MailUser -Identity $dn -Confirm:$False
			}
			
			$student.memberof | Where-Object { $_ -ne $googleAppDn } | ForEach-Object { Remove-UserGroup $student.distinguishedName $_ }
			
		}
		
		if($student.distinguishedName)
		{
			<$userControl = [System.Convert]::ToInt32($student.userAccountControl.ToString())
			$description = "IT - Not Coming, {0}" -f $student.description.ToString()
			
			Write-Host "`tUpdate account information"
			$student.userAccountControl = ($userControl -bor $disableAccountFlag)
			$student.description = $description
			$student.CommitChanges()
			
			foreach($grp in $admitGroups)
			{
				Remove-UserGroup $student.distinguishedName $grp
			}
			Write-Host "`tMove to admitted - not coming OU"
			$student.PSBase.moveto("LDAP://" + $admittedNotComingOu)
		}
	}
	
}



function Move-Account
{
    param(
        $student = $(throw "student is required"), 
        $groupToRemove, 
        $groupToAdd,
        $studentDn = "OU=Students,OU=Network Users,DC=lawnet,DC=lcl"
    )

    Write-Host "Move user ($($student.cn))"
    
    $groupToRemove | ForEach-Object { Remove-UserGroup $student.distinguishedName $_ }
    $groupToAdd | ForEach-Object { Add-UserGroup $student.distinguishedName $_ }    
 
    $inheritanceFlags = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
    $propagateFlags = [System.Security.AccessControl.PropagationFlags]"None"
    $folderRule = New-Object System.Security.AccessControl.FileSystemAccessRule($student.userPrincipalName, "Modify", $inheritanceFlags, $propagateFlags, "Allow")
       
    if(Test-Path $student.homeDirectory)
    {
        Write-Host "`tHome Folder" $student.homeDirectory "already exists"
    }
    else
    {
        Write-Host "`tCreate Home Folder" $student.homeDirectory
        New-Item $student.homeDirectory -type directory | out-null
        #New-Item ($student.homeDirectory.ToString() + "\Document") -type directory | out-null
        #New-Item ($student.homeDirectory.ToString() + "\Document\Templates") -type directory | out-null
    }
    
    Write-Host "`tSet Home Folder Permission"
    
    $acl = Get-Acl $student.homeDirectory
    $acl.SetAccessRule($folderRule)
    Set-Acl -Path $student.homeDirectory -AclObject $acl
        
    if(!$student.legacyExchangeDN)
    {
        $dn = $student.distinguishedName.ToString()
        $alias = $student.sAMAccountName.ToString()
        $addr = $student.targetAddress.ToString()
        $dc = $student.psbase.Options.GetCurrentServerName().ToString()
        
        if($addr -eq "")
        {
            $addr = "SMTP:" + $student.mail.ToString()
        
            #set targetAddress if we have a valid mail field
            if($addr -ne "SMTP:")
            {    
                $student.put("targetAddress", $addr)
                $student.SetInfo()
            }
        }
        
        #make sure we have a valid target address
        if($addr -ne "SMTP:")
        {
            Write-Host "`tMail enable user ($addr)"
            Enable-MailUser -Identity $dn -Alias $alias -ExternalEmailAddress $addr -DomainController $dc | out-null
        }
        else
        {
            Write-Host "`tAccount has no email address set"
        }
    }
    
    #do this last to avoid any issues with DC sync issue
    $studentOu = [ADSI]("LDAP://$studentDn")
    Write-Host "`tMove to OU ($($studentOu.name))"
    $student.PSBase.moveto($studentOu)
}

function checkCommonName
{
    param (
           [string] $CommonName = $(throw"CommonName is required")
    )
    
    #Use system.directorysearcher object to do LDAP query
    $strFilter = "(sAMAccountName=$CommonName)"
    $objDomain = New-Object System.DirectoryServices.DirectoryEntry #("LDAP://dc=lawnet, dc=lcl")
    $objSearcher = New-Object System.DirectoryServices.DirectorySearcher
    $objSearcher.SearchRoot = $objDomain
    $objSearcher.Filter = $strFilter
    $objSearcher.SearchScope = "Subtree"
    
    if ($objSearcher.FindOne() -ne $null){ 
        return $false    
    }
    else{
        return $true
    }
}

function Add-NewAdmit
{
	param(
                   $data = $(throw "data is required."), 
                   $admitGroup = $(throw "admit group is required."),
                   $admitDn = "OU=Admitted Students,OU=Network Users,DC=lawnet,DC=lcl",
                   $studentGroup = $(throw "student group is required."),
                   $studentDn = "OU=Students,OU=Network Users,DC=lawnet,DC=lcl",
        [string]   $homeRoot = "\\uclawdata1\StuHome\",
        #[string]  $googleAppDn = "CN=Google Apps Users,OU=Network Users,DC=lawnet,DC=lcl",                   
        [DateTime] $moveStart = "8/1/" + [DateTime]::Now.Year.ToString(),
        [DateTime] $moveEnd = "9/1/" + [DateTime]::Now.Year.ToString(),
        [string]   $upnDomain = "lawnet.lcl",
        [string]   $emailDomain = "lawnet.ucla.edu",
        [int]$createdCount
    )
    
    $disableAccountFlag = 2
    $pwNotExpireFlag = 65536 #0x10000
    $expiredate = [DateTime]::Now.Date.ToFileTimeUtc().ToString()

    #Get default culture for text formatting
    $textInfo = (Get-Culture).TextInfo
    
    $admitOu = [ADSI]("LDAP://$admitDn")
    
    [int]$Year = getYear($data.Username)
    [string]$Username = GetUserName $data.Username $data.LastName $data.FirstName $data.MiddleName $Year
    if(!$Username){
        Write-Error "Error with GetUserName. Using OG."
        $Username = $data.Username
    }
    Write-Host "`r"
    Write-Host "`rCreate user ($Username)"
    Write-Host "`r"
	
    if (checkCommonName ($data.CommonName)){
	   $student = $admitOu.Create("user", "cn=$($data.CommonName)")
    }
    
    if($student)
    {
        $student.put("sAMAccountName", $textInfo.ToUpper($Username))
        
        if($data.FirstName.Length -gt 0)
        {
            $student.put("givenName", $textInfo.ToTitleCase($data.FirstName.ToLower()))
        }
        
        if($data.LastName.Length -gt 0)
        {
            $student.put("sn", $textInfo.ToTitleCase($data.LastName.ToLower()))
        }
        
        if($data.MiddleName.Length -gt 0)
        {
            $student.put("middleName", $textInfo.ToTitleCase($data.MiddleName.ToLower()))
        }        
        
        $student.put("userPrincipalName", "$Username@$upnDomain")
        $student.put("displayName", "$($data.CommonName)")
        $student.put("name", "$($data.CommonName)")
        $student.put("homeDrive", "H:")
        $student.put("homeDirectory", "$homeRoot$($Username)")
        $student.put("description", "$($data.Description)")
        $student.put("title", "$($data.Description)")
        $student.put("employeeID", "$($data.StudentId)")
        $student.put("targetAddress", "SMTP:$($textInfo.ToLower($Username))@$emailDomain")
        $student.put("mail", "$($textInfo.ToLower($Username))@$emailDomain")
        $student.put("extensionAttribute2", "$($data.BirthDate)")
        $student.put("accountExpires", "$expiredate")
        
        if($data.SocialSecurityLastFour.Length -gt 1 -and $data.SocialSecurityLastFour -ne "NULL")
        {
            $student.put("extensionAttribute3", "$($data.SocialSecurityLastFour)")
        }
        
        $student.SetInfo()
        
        $userFlag = $student.get("userAccountControl")
        $userFlag = ($userFlag -band (-bnot $disableAccountFlag)) -bor $pwNotExpireFlag
        $student.put("userAccountControl", $userFlag)
        $student.SetInfo()
        
        #set password
        $student.psbase.invoke("SetPassword", $data.Password)
        $student.psbase.CommitChanges()
        
        #add to admitted student group(s)
        $admitGroup | ForEach-Object { Add-UserGroup $student.distinguishedName $_ }
        
        $global:createdCount++
        $global:usersCreated +="`r $($student.displayName), $Username"

        <#see if we need to migrate to student OU - Is currently being done in NewAdmit script directly
        if($moveStart -le [DateTime]::Now -and $moveEnd -gt [DateTime]::Now)
        {
            Move-Account $student $admitGroup $studentGroup $studentDn
        }#>
    } #>
    else
    {
        $ErrMsg = "`r" + "Unable to create user " + ($data.CommonName)
        $global:failCount++
        $global:usersFail += "`r $($student.displayName), $Username"
        Write-Error $ErrMsg
    }

}

function Add-SIR
{
    param(
		$data = $(throw "data is required."), 
        $admitDn = "OU=Admitted Students,OU=Network Users,DC=lawnet,DC=lcl",
		[string] $googleAppDn = "CN=Google Apps Users,OU=Network Users,DC=lawnet,DC=lcl"
	)
	$admitOu = [ADSI]("LDAP://$admitDn")
	$student = [ADSI]("LDAP://CN=$data,$admitDn")
    
	if($student){
	    Add-UserGroup $student.distinguishedName $googleAppDn
	}
    
	#param(
	#	$data = $(throw "data is required."), 
    #    $admitDn = "OU=Admitted Students,OU=Network Users,DC=lawnet,DC=lcl",
	#	[string] $googleAppDn = "CN=Google Apps Users,OU=Network Users,DC=lawnet,DC=lcl"
	#)
	#$admitOu = [ADSI]("LDAP://$admitDn")
	#$student = [ADSI]("LDAP://CN=$($data.CommonName),$admitDn")
	#if($student){
	#	Write-Host "`rAdd user to Google groups ($($data.CommonName))"
	#		$student = $result.GetDirectoryEntry()
	#		Add-UserGroup $student.distinguishedName $googleAppDn
	#}
}

function startLog
{
    param ($path, $now)
    
    #$error.clear()
    start-transcript $path
    Write-Host "`r"
    Write-Host "`tTransaction Log for $path on: "$now "`r"
    Write-Host "`r" 
    
}

function writeErrors
{
    if($?){
        Write-Host "`r"
        Write-Host "`r"
        Write-Host "`r `tNo Errors!`r"
        Write-Host "`r"
        return $true
    } 
    else{
        Write-Host "`r"
        Write-Host "`r"
        Write-Host "`r `t ERRORS occurred!`r"
        Write-Host "`r"

       return $false
    }
}

function sendLog
{
    param(
        [string]$path,
        [string]$now,
        $err
    )
    
    if ($err){
        $filecontents = Get-Content $path 
        $server = [System.Net.Dns]::GetHostName()
        Email "telesino@law.ucla.edu, myersd@law.ucla.edu, SANCHEZM@law.ucla.edu, lawnetdevelopment@law.ucla.edu, systems@law.ucla.edu" "No errors. Transaction Log for NewAdmit on: $now" "Log can be found here on $server $path.`r `rNumber of accounts created: $createdCount. $usersCreated `r `rNumber of account failures: $failCount. $Usersfail"
    }

    else{
        $filecontents = Get-Content $path 
        $server = [System.Net.Dns]::GetHostName()
        Email "myersd@law.ucla.edu, SANCHEZM@law.ucla.edu, lawnetdevelopment@law.ucla.edu, systems@law.ucla.edu" "THERE WAS AN ERROR: Transaction Log for NewAdmit on: $now" "Log can be found here on $server $path.`r `rNumber of accounts created: $createdCount.`r `rNumber of account failures: $failCount.`r `rLog contents here:`r`r$filecontents"
    }
}

function sendErrLog
{
    param(
        [string]$path,
        [string]$now,
        $err
    )
    
    if (!$err){
        $filecontents = Get-Content $path 
        $server = [System.Net.Dns]::GetHostName()
        Email "lawnetdevelopment@law.ucla.edu" "THERE WAS AN ERROR: Transaction Log for Add-SIRGoogle on: $now" "Log can be found here on $server $path.`r `rLog contents here:`r`r$filecontents"
    }
}
	


<#original code prior to SIR NewAdmit differentiation as of 9/2012
function Add-Account
{
    param(
                   $data = $(throw "data is required."), 
                   $admitGroup = $(throw "admit group is required."),
                   $admitDn = "OU=Admitted Students,OU=Network Users,DC=lawnet,DC=lcl",
                   $studentGroup = $(throw "student group is required."),
                   $studentDn = "OU=Students,OU=Network Users,DC=lawnet,DC=lcl",
        [string]   $homeRoot = "\\uclawdata1\StuHome\",
        [string]   $googleAppDn = "CN=Google Apps Users,OU=Network Users,DC=lawnet,DC=lcl",                   
        [DateTime] $moveStart = "8/1/" + [DateTime]::Now.Year.ToString(),
        [DateTime] $moveEnd = "9/1/" + [DateTime]::Now.Year.ToString(),
        [string]   $upnDomain = "lawnet.lcl",
        [string]   $emailDomain = "lawnet.ucla.edu"
    )

    $disableAccountFlag = 2
    $pwNotExpireFlag = 65536 #0x10000
    $expiredate = [DateTime]::Now.Date.ToFileTimeUtc().ToString()

    #Get default culture for text formatting
    $textInfo = (Get-Culture).TextInfo
    
    $admitOu = [ADSI]("LDAP://$admitDn")

    Write-Host "Create user ($($data.CommonName))"
	
	$student = $admitOu.Create("user", "cn=$($data.CommonName)")
    
    if($student)
    {
        $student.put("sAMAccountName", $textInfo.ToUpper($data.Username))
        
        if($data.FirstName.Length -gt 0)
        {
            $student.put("givenName", $textInfo.ToTitleCase($data.FirstName.ToLower()))
        }
        
        if($data.LastName.Length -gt 0)
        {
            $student.put("sn", $textInfo.ToTitleCase($data.LastName.ToLower()))
        }
        
        if($data.MiddleName.Length -gt 0)
        {
            $student.put("middleName", $textInfo.ToTitleCase($data.MiddleName.ToLower()))
        }        
    
        $student.put("userPrincipalName", "$($data.Username)@$upnDomain")
        $student.put("displayName", "$($data.CommonName)")
        $student.put("name", "$($data.CommonName)")
        $student.put("homeDrive", "H:")
        $student.put("homeDirectory", "$homeRoot$($data.UserName)")
        $student.put("description", "$($data.Description)")
        $student.put("title", "$($data.Description)")
        $student.put("employeeID", "$($data.StudentId)")
        $student.put("targetAddress", "SMTP:$($textInfo.ToLower($data.Username))@$emailDomain")
        $student.put("mail", "$($textInfo.ToLower($data.Username))@$emailDomain")
        $student.put("extensionAttribute2", "$($data.BirthDate)")
        $student.put("accountExpires", "$expiredate")
        
        if($data.SocialSecurityLastFour.Length -gt 0 -and $data.SocialSecurityLastFour -ne "NULL")
        {
            $student.put("extensionAttribute3", "$($data.SocialSecurityLastFour)")
        }
        
        $student.SetInfo()
        
        $userFlag = $student.get("userAccountControl")
        $userFlag = ($userFlag -band (-bnot $disableAccountFlag)) -bor $pwNotExpireFlag
        $student.put("userAccountControl", $userFlag)
        $student.SetInfo()
        
        #set password
        $student.psbase.invoke("SetPassword", $data.Password)
        $student.psbase.CommitChanges()
        
        #add to admitted student group(s)
        $admitGroup | ForEach-Object { Add-UserGroup $student.distinguishedName $_ }
        
        #add to google app group
        Add-UserGroup $student.distinguishedName $googleAppDn           

        #see if we need to migrate to student OU
        if($moveStart -le [DateTime]::Now -and $moveEnd -gt [DateTime]::Now)
        {
            Move-Account $student $admitGroup $studentGroup $studentDn
        }
    }
    else
    {
        Write-Host "Unable to create user ($(data.CommonName))"
    }
}#>