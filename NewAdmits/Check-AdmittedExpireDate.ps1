$admitOU = [ADSI]"LDAP://OU=Admitted Students,OU=Network Users,DC=lawnet,DC=lcl"

$count = 0

Write-Host "Active accounts:`r"
$admitOU.children | ? {$_.AccountExpirationDate.Day -lt 1} | % { 
[string]$desc = $_.description
if($desc.Contains("IT")){}
else {Write-Host $_.cn 
$count++}}

Write-Host "`r"

Write-Host "Expired Accounts:`r"
$admitOU.children | ? {$_.AccountExpirationDate.Day } | % { 
[string]$desc = $_.description
if($desc.Contains("IT")){}
else {Write-Host $_.cn  
$count++} }

Write-Host "Number of total Admits: $count"

<#$now = [DateTime]::Now.Date.ToString("MM.dd.yyyy")
$scriptpath = $MyInvocation.MyCommand.Path
$path = (Split-Path $scriptpath) + "\logs\NewAdmitLog - $now.txt"

Write-Host $path

$now = [DateTime]::Now.Date.ToString("MM.dd.yyyy")
$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
Write-host "My directory is $dir"
$path = "logs\NewAdmitLog - $now.txt"
$path = Join-Path -path $dir -childpath $path

Write-Host $path#>
    
    <#if( $_.AccountExpirationDate ){
            Write-Host $_.cn $_.AccountExpirationDate 
    }#if !AccountExpirationDate -> accountexpires -eq 0
}
#>
        