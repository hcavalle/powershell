#4647 = logoff event
function Process-Logoff{
    
    param(
        [string]$labMachine,
        $cmd,
        [DateTime]$lastlogon
    )
    [DateTime]$date = (Get-Date)
    #$curdate = $date.ToString("s")
    #$curdate += 'Z'
    #$StartTime =  $lastlogon.ToString("s")
    #$StartTime += 'Z'
    $lastLogonTime =  $lastlogon.ToString("s")
    $lastLogonTime += 'Z'
    $lastLogonTime
    $milisecs = $date - $lastlogon 
    #[double]$diff = $milisecs.TotalMilliseconds
    $diff
    
    $ErrorActionPreference = "Stop" 
    #$xmlquery = "<QueryList><Query Id=""0"" Path=""Security""><Select Path=""Security"">*[System[(EventID=4647) and TimeCreated[timediff('$StartTime', '$curdate')]]] </Select></Query></QueryList>" (@SystemTime, '2013-02-10T06:30:00Z') &lt;= 43200000]]]
    $xmlquery = [string]::Format("<QueryList><Query Id=""0"" Path=""Security""><Select Path=""Security"">*[System[(EventID=4647) ]] </Select></Query></QueryList>", $lastLogonTime)
    try{$Entries = Get-WinEvent -FilterXml $xmlquery -ComputerName $labMachine | ? { $_.TimeCreated -gt $lastLogonTime}}
    
    catch { 
        $global:failures += "`r Error getting logoff events for $labMachine. Or there were no entries."
        $error.clear()
        $Entries = $null
    }
    if($Entries){
        #bind params to var
        if($cmd){
            $timeparam = $cmd.Parameters.Add("@logofftime", [System.Data.SqlDbType]"DateTime") #create query
            $userparam = $cmd.Parameters.Add("@username", [System.Data.SqlDbType]"Char", 16) #create query
            $computerparam = $cmd.Parameters.Add("@computername", [System.Data.SqlDbType]"Char", 9) #create query
            $logonidparam = $cmd.Parameters.Add("@logonid", [System.Data.SqlDbType]"varchar") #create query
        }
    
        foreach($entry in $Entries){
           try {$xmlstring = $entry.ToXml()}
           catch{
                $err = $error[0]
                $global:failures += "`r Converting event logs to xml failed $labmachine"
                $error.clear()
                break     
           }
           $xmlEntry = New-Object System.Xml.XmlDocument
           $xmlEntry.LoadXML($xmlString) 
           [int]$event = $xmlEntry.Event.System.EventID
           [DateTime]$time = $xmlEntry.Event.System.TimeCreated.SystemTime
           
           $eventdata = $xmlEntry.Event.EventData 
           $processdata = $eventdata.ChildNodes | ? {$_.Name -eq "ProcessName"}
           $processname = $processdata.innertext
           $userdata = $eventdata.ChildNodes | ? { $_.Name -eq "TargetUserName" }
           $username = $userdata.innertext
           $logondata= $eventdata.ChildNodes | ? { $_.Name -eq "TargetLogonId" }
           $logonid = $logondata.innertext
           $computername = $xmlEntry.Event.System.Computer
           
           # assign value to params
           if($cmd){
               $timeparam.Value = $time
               $userparam.Value = $username
               $computerparam.Value = $computername
               $logonidparam.Value = $logonid
           }
           
           Write-Host "Logoff " $computername $username $time $logonid
           $var = $cmd.ExecuteNonQuery() 
        }
    }
}


#4648 = logon event
function Process-Logon {
    param(
        [string]$labMachine,
        $cmd,
        [DateTime]$lastlogon
    )
    [DateTime]$date = (Get-Date)
    #$curdate = $date.ToString("s")
    #$curdate += 'Z'
    #$StartTime =  $lastlogon.ToString("s")
    #$StartTime += 'Z'
    $lastLogonTime =  $lastlogon.ToString("s")
    $lastLogonTime += 'Z'
    
    $ErrorActionPreference = "Stop"
    $xmlquery = [string]::Format("<QueryList><Query Id=""0"" Path=""Security""><Select Path=""Security"">*[System[(EventID=4624)]] and *[EventData[Data[@Name='LogonType']=2]] </Select></Query></QueryList>", $lastLogonTime)
    try {$Entries = Get-WinEvent -FilterXml $xmlquery -ComputerName $labMachine | ? { $_.TimeCreated -gt $lastLogonTime} }
    
   catch { 
        $err = $error[0]
        $global:failures += "`r Error getting logon events for $labmachine"
        $error.clear()
        $Entries = $null
    } 
    
    if($Entries){
        
        #bind params to var
        if($cmd){
            #$eventparam = $cmd.Parameters.Add("@event", [System.Data.SqlDbType]"Int") #create query
            $timeparam = $cmd.Parameters.Add("@logontime", [System.Data.SqlDbType]"DateTime") #create query
            $userparam = $cmd.Parameters.Add("@username", [System.Data.SqlDbType]"varchar") #create query
            $computerparam = $cmd.Parameters.Add("@computername", [System.Data.SqlDbType]"Char", 9) #create query
            $logonidparam = $cmd.Parameters.Add("@logonid", [System.Data.SqlDbType]"varchar") #create query
        }
        
        foreach($entry in $Entries){
           try {$xmlstring = $entry.ToXml()}
           catch{
                $err = $error[0]
                $global:failures += "`r Converting event logs to xml failed $labmachine"
                $error.clear()
                break
           }
           $xmlEntry = New-Object System.Xml.XmlDocument
           $xmlEntry.LoadXML($xmlString) 
           [int]$event = $xmlEntry.Event.System.EventID
           [DateTime]$time = $xmlEntry.Event.System.TimeCreated.SystemTime
           
           $eventdata = $xmlEntry.Event.EventData   
           $processdata = $eventdata.ChildNodes | ? {$_.Name -eq "ProcessName"}
           $processname = $processdata.innertext
           $userdata = $eventdata.ChildNodes | ? { $_.Name -eq "TargetUserName" }
           $username = $userdata.innertext
           $logondata= $eventdata.ChildNodes | ? { $_.Name -eq "TargetLogonId" }
           $logonid = $logondata.innertext
           $computername = $xmlEntry.Event.System.Computer
           
           # assign value to params
           if($cmd){
               $timeparam.Value = $time
               $userparam.Value = $username
               $computerparam.Value = $computername
               $logonidparam.Value = $logonid
           }
          
           Write-Host "Logon" $username $computername $time $logonid
           $var = $cmd.ExecuteNonQuery() 
        }
    }
}

###################LOGGING CODE ######################
function Email{

    param(
        [string]$emailto,
        [string]$subject,
        [string]$body
    )

    $emailFrom = "labstats@law.ucla.edu"
    $smtpServer = "smtp.ucla.edu"
    $smtp = new-object Net.Mail.SmtpClient($smtpServer)
    $smtp.Send($emailFrom, $emailTo, $subject, $body)

}
function startLog
{
    param ($path, $now)
    
    start-transcript $path
    Write-Host "`r"
    Write-Host "Transaction Log for $path on: "$now "`r"
    Write-Host "`r" 
    
}

function writeErrors
{
    if($global:failures.length -eq 0){
        Write-Host "`r"
        Write-Host "`r"
        Write-Host "`r No Errors!`r"
        Write-Host "`r"
        return $true
    } 
    else{
        Write-Host "`r"
        Write-Host "`r"
        Write-Host "`r ERRORS occurred!`r"
        Write-Host  $global:failures
        Write-Host "`r`r"

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
        Email "harrison@law.ucla.edu" "No errors. Transaction Log for NewAdmit on: $now" "Log can be found here on $server $path."
    }

    else{
        $filecontents = Get-Content $path
        $server = [System.Net.Dns]::GetHostName()
        Email "harrison@law.ucla.edu" "THERE WAS AN ERROR: Log for LabStats: $now" "Log can be found here on $server $path. `r Errors are: `r `r $global:failures"
    }
}



<#function Load-Module
{
    param(
        [string] $scriptPath
    )
    
    if((!$scriptPath) -or $scriptPath -eq "")
    {
        $scriptLoc = Split-Path -Parent $MyInvocation.ScriptName
    }
    else
    {
        $scriptLoc = $scriptPath
    }
    
    if(!$scriptLoc.EndsWith("\"))
    {
        $scriptFile = ($scriptLoc + "\LabStats.psm1")
    }
    else
    {
        $scriptFile = ($scriptLoc + "LabStats.psm1")
    }
    
    if(Test-Path -LiteralPath $scriptFile -PathType Leaf)
    {  
        if(Get-Module | Where-Object { $_.Name -eq "LabStats" })
        {
            Remove-Module -Name "LabStats"
        }
        
        Import-Module -Name $scriptFile
    }
    else
    {
        throw ("Unable to locate script module " + $scriptFile)
    }
    
    if(!(Get-Command | Where-Object {$_.Name -eq "Enable-MailUser"}))
    {
        ImportSystemModules
    }
}

#Process-Lab-Machines 
$now = [DateTime]::Now.Date.ToString("yyyy.MM.dd")
$scriptpath = $MyInvocation.MyCommand.Path
$path = (Split-Path $scriptpath) + "\logs\LabStatsLog - $now.txt"
[string]$global:failures = ""

#startLog $path $now
#

#logon table connection
$connStr = "data source=UCLAWDEV3;initial catalog=General;Integrated Security=true;"
$conn = New-Object System.Data.SqlClient.SqlConnection($connStr)
$conn.open()
 
$cmdStr ="INSERT INTO [General].[dbo].Lab_Logon (logontime,username,computername,logonid) VALUES (@logontime, @username, @computername,@logonid)" 
$cmdStr2 = "UPDATE [General].[dbo].Lab_Logon SET logofftime=@logofftime WHERE logonid=@logonid" 

<########### Run Query on all lab machines ###########
[string[]]$labsOU = "OU=Labs - Win 7,OU=Network Resources,DC=lawnet,DC=lcl", "OU=Clinical,OU=Network Resources,DC=lawnet,DC=lcl"
foreach($ou in $labsOU){
    $labmachineCollection = [ADSI]("LDAP://$ou")
    
    #$labsOU = "OU=Labs - Win 7,OU=Network Resources,DC=lawnet,DC=lcl"
    #$labmachineCollection = [ADSI]("LDAP://$labsOU")
    
    foreach($machine in $labmachineCollection.children) {
        $labMachine = $machine.cn
        $bool = $TRUE;
        [string[]]$exclude = "stulabfrontdesk", "mootcourt7", "BLSA7", "APALJ7"
        foreach ($ex in $exclude){ if ($ex -eq $labmachine) {$bool = $FALSE}}
        
        if($bool){
           #get last logon entry, to knwo where to start. 
            $getmax= getMax
            
            $cmd = New-Object System.Data.SqlClient.SqlCommand($cmdStr, $conn)
            Process-Logon $labMachine $cmd $max
            
            #repeat get max with logoff
            $cmd2 = New-Object System.Data.SqlClient.SqlCommand($cmdStr2, $conn)
            Process-Logoff $labMachine $cmd2 $max
            break
         }
    } 
} #

##############Test single machine 
$labMachine =  "stulab137" #"stulab132" #"lawnet7"
$cmd = New-Object System.Data.SqlClient.SqlCommand($cmdStr, $conn)
$cmd2 = New-Object System.Data.SqlClient.SqlCommand($cmdStr2, $conn)

#get last logon entry, to knwo where to start. 
$max= getMax $conn
Process-Logon $labMachine $cmd $max
Process-Logoff $labMachine $cmd2 $max
#

##### run stored proc on sql to calc session length ##########
$sproccmd = New-Object System.Data.SqlClient.SqlCommand("exec dbo.calcLogonLength",$conn)
try{$sproccmd.ExecuteNonQuery()}
catch{Write-Host "Failed stored procedure: " $error[0]}
#############################

$conn.close()

#stop-transcript

$err = writeErrors
Write-Host "Test: " $global:failures

try {sendLog $path $now $err}
catch{"error sending log."}
#>
