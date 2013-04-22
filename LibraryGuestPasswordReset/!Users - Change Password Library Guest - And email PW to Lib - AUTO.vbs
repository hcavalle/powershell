rem For Student user administrators to change passwords on accounts in the Student OU

rem declar vars
Dim strPw 
Dim sTargetOU
Dim pwBaseLength
pwBaseLength = 10
Dim pwBase(10)
pwBase(0) = "UCLA_"
pwBase(1) = "BruinS"
pwBase(2) = "uclaLaw"
pwBase(3) = "Lawlib"
pwBase(4) = "HilGard"
pwBase(5) = "YoungDr"
pwBase(6) = "WalBil"
pwBase(7) = "CircDesk"
pwBase(8) = "Hammer"
pwBase(9) = "Wilshire"

rem generate random number to select basePW
Randomize
intRand = Int((0 - pwBaseLength) * Rnd) + 10
strPw = pwBase(intRand)

rem ###########OLD CODE
rem code for PW generator
rem intUpperLimit = 4
rem intLowerLimit = 4
 
rem Randomize  
rem intCharacters = Int(((intUpperLimit - intLowerLimit + 1) * Rnd) _
rem     + intUpperLimit) 
rem #########end old code 

intRandCharLimit = 4 rem determines how man random characters are appended to the end of basePW
 
rem these vars determine with range of ASCII table random chars are pulled from
intUpperLimit = 64
intLowerLimit = 33
 
For i = 1 to intRandCharLimit
    Randomize
    intASCIIValue = Int(((intUpperLimit - intLowerLimit + 1) * Rnd) _
        + intLowerLimit)  
    strPw = strPw & Chr(intASCIIValue)
Next
 
rem ####### FOR TESTING Wscript.Echo strPw

rem ########FOR TESTING MsgBox strPw

rem code for setting password on lib guest OU 
rem OLD CODE strPw = InputBox("Enter the new password for all LIBGUEST users") rem for manual pw

 Set oTargetOU = GetObject("LDAP://ou=libguest,ou=network users,dc=lawnet,dc=lcl")

 oTargetOU.Filter = Array("user")

 For each usr in oTargetOU

	 usr.setpassword strPw

 Next

rem code for emailing it to circ desk
Set objMessage = CreateObject("CDO.Message") 
objMessage.Subject = "Library guest passwords" 
objMessage.From = "help@law.ucla.edu" 
rem objMessage.To = "harrison@law.ucla.edu; " rem for testing
objMessage.To = "circulation@law.ucla.edu; jason@law.ucla.edu; andres@law.ucla.edu; reid@law.ucla.edu; daniel@law.ucla.edu" 
objMessage.TextBody = "I have changed the password for the guest accounts to: "& vbCrLf & vbCrLf & strPw & vbCrLf & vbCrLf & "Please distribute this password as needed."

'==This section provides the configuration information for the remote SMTP server.
'==Normally you will only change the server name or IP.
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

'Name or IP of Remote SMTP Server
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.ucla.edu"

'Server port (typically 25)
objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = false

objMessage.Configuration.Fields.Item _
("http://schemas.microsoft.com/cdo/configuration/sendusername") = "help@law.ucla.edu"

rem objMessage.Configuration.Fields.Item _
rem ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""


objMessage.Configuration.Fields.Update

'==End remote SMTP server configuration section==

objMessage.Send

rem ####### For error catching
If Err.Number <> 0 Then
    ' VB exception handling example
    Set objMessage = CreateObject("CDO.Message") 
	objMessage.Subject = "Library guest passwords" 
	objMessage.From = "jason@law.ucla.edu" 
	objMessage.To = "harrison@law.ucla.edu; " rem for testing
	rem objMessage.To = "circulation@law.ucla.edu; jason@law.ucla.edu" 
	objMessage.TextBody = "There was an error setting the library guest password. Set manually using the Manual script and investigate cause."

	'==This section provides the configuration information for the remote SMTP server.
	'==Normally you will only change the server name or IP.
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 

	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.ucla.edu"

	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = false

	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusername") = "help@law.ucla.edu"
	
	rem objMessage.Configuration.Fields.Item _
	rem ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = ""
        
End If

rem ############## for interactive mode/testing MsgBox "Finished"
