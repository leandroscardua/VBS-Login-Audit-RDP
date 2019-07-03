
firstname = GetUserFirstname()
username = GetUsername()
email = GetEmail()
computer = GetComputer()
sourcecomputer = GetSourceComputer()

loginReason = ""
Do While (loginReason = "")
    loginReason = InputBox("Hi " + firstname + ", please describe the reason of your login:", "Login Audit")
Loop

eventDescription = "ComputerName: " & sourcecomputer & Chr(13)& "UserName: " & username & Chr(13)& "ServerName: " & computer & Chr(13)& "Reason: " & loginReason & Chr(13) 
' Change the /T the type of Event /ID change the id number /SO Change the name or source of the event.
Set WshShell = WScript.CreateObject("WScript.Shell")
strCommand = "eventcreate /T Information /ID 100 /L Application /SO LoginAudit /D " & _
    Chr(34) & eventDescription & Chr(34)
WshShell.Run strcommand

Function GetUserFirstname()
    Set objSysInfo = CreateObject("ADSystemInfo")
    Set objCurrentUser = GetObject("LDAP://" & objSysInfo.UserName)
    GetUserFirstname = objCurrentUser.givenName
End Function

Function GetUsername()
    Set objNetwork = CreateObject("Wscript.Network")
    GetUsername = objNetwork.UserName
End Function

Function GetEmail()
   Set objSysInfo = CreateObject("ADSystemInfo")
   Set objCurrentUser = GetObject("LDAP://" & objSysInfo.UserName)
   GetEmail = objCurrentUser.mail
End Function

Function GetComputer()
	Set objSysInfo = CreateObject( "WinNTSystemInfo" )
	strComputerName = objSysInfo.ComputerName
    GetComputer = strComputerName
End Function

Function GetSourceComputer()
	Set Sh = CreateObject("WScript.Shell")
    sys = Sh.ExpandEnvironmentStrings("%CLIENTNAME%")
	GetSourceComputer = sys
End Function

' Change the email from, CC and HTMLBODY
 Set objMessage = CreateObject("CDO.Message") 
 objMessage.Subject = "Login Audit" 
 objMessage.From = "noreply@example.com" 
 objMessage.cc = "noreply@example.com" 
 objMessage.To = email
 objMessage.HTMLBody =  "<p style=""text-align: center;""><span style=""text-decoration: underline; background-color: #ffff00; font-size: 20pt;""><strong> Remote Desktop Access - RDP</strong></span></p> <p style=""text-align: center;""><strong><span style=""font-size: 12.0pt; line-height: 107%;"">*** REPLY to this email if you have any question ***</span></strong></p> <p> Dear <strong><u><span style=""color: red;""> " & firstname & " </span></u></strong>,</p> Your account was successfully logged on the server <strong><u> " & computer & "</u></strong> from <strong><u>" & sourcecomputer & " </u></strong>  <p> Providing the following reason: <strong>" & loginReason & "</strong> </p> <p><strong><span style=""font-size: 12.0pt; line-height: 107%;""> Please contact immediately the admin if this is an unauthorized access.</span></strong></p>"

 Set objConfig = objMessage.Configuration
 objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
 ' Add teh IP of the SMTP Server
 objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "XXX.XXX.XXX.XXX"
 ' Change the Port if necessary
 objConfig.Fields("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25  
 objConfig.Fields.Update
 objMessage.Send

Set objShell = WScript.CreateObject("WScript.Shell")
objShell.Run("explorer")
Set objShell = Nothing
