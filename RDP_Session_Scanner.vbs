/* Originally written in November 2012 */

'Setup for variables

MySMTPServer = "MySMTPServer"
MyMessageSubject = "RDP Logins - My Active Directory Domain Name"
MyMessageSendTo = "MyEmail@MyEmailDomain.COM"

UserListing = ""
DeadServers = "<hr>Servers in Active Directory that do not respond to PING:<br>"
DeadServerCount = 0

BadWMI = "<hr>Servers where WMI is broken:<br>"
BadWMICount = 0

LoggedInCounter = 0


' Setup ADO objects.
Set adoCommand = CreateObject("ADODB.Command")
Set adoConnection = CreateObject("ADODB.Connection")
adoConnection.Provider = "ADsDSOObject"
adoConnection.Open "Active Directory Provider"
adoCommand.ActiveConnection = adoConnection

' Search entire Active Directory domain.
Set objRootDSE = GetObject("LDAP://RootDSE")
strDNSDomain = objRootDSE.Get("defaultNamingContext")
strBase = "<LDAP://" & strDNSDomain & ">"

' Filter on user objects.
strFilter = "(&(objectCategory=computer)(operatingSystem=*server*))"

' Comma delimited list of attribute values to retrieve.
strAttributes = "name"

' Construct the LDAP syntax query.
strQuery = strBase & ";" & strFilter & ";" & strAttributes & ";subtree"
adoCommand.CommandText = strQuery
adoCommand.Properties("Page Size") = 100
adoCommand.Properties("Timeout") = 30
adoCommand.Properties("Cache Results") = False

' Run the query.
Set adoRecordset = adoCommand.Execute

' Enumerate the resulting recordset.
Do Until adoRecordset.EOF
	'On Error Resume Next
	strHostname = adoRecordset.Fields("name").Value
	If CheckStatus(strHostname) = False Then
		DeadServers = DeadServers & strHostName & "<br>"
	Else
		ExamineSessions(strhostname)
	End If
	
	'Move to the next record in the recordset.
    adoRecordset.MoveNext
Loop

If LoggedInCounter > 0 then
	MailMessage = "Listing of People still logged into Servers:<br> <br>"
	MailMessage = MailMessage & "<table border=1>"
	MailMessage = MailMessage & "<tr><td>Server</td><td>User</td><td>Logged on Since</td></tr>"
	MailMessage = MailMessage & userlisting
	MailMessage = MailMessage & "</table>"
Else
	MailMessage = ""
End If

if BadWMICount > 0 then
	MailMessage = MailMessage & BadWMI
end if

if DeadServerCount > 0 then
	MailMessage = MailMessage & DeadServers
end if


Sendmail MyMessageSendTo, MyMessageSubject


' Clean up.
adoRecordset.Close
adoConnection.Close

Set adoRecordset = Nothing
Set objRootDSE = Nothing
Set adoConnection = Nothing
Set adoCommand = Nothing

Function ExamineSessions(ServerName)
on error resume next
Set objWMIService = GetObject("winmgmts:\\" & ServerName & "\root\cimv2")
Set colProcessList = objWMIService.ExecQuery("Select * from Win32_Process Where Name='explorer.exe'")

if isnull(colProcessList) = False then

	For Each objProcess in colProcessList

	objProcess.GetOwner strNameOfUser, strUserDomain
	strOwner = strUserDomain & "\" & strNameOfUser

	Startdate = CDate(Mid(objprocess.creationdate, 5, 2) & "/" & _
	     Mid(objprocess.creationdate, 7, 2) & "/" & Left(objprocess.creationdate, 4) _
        	 & " " & Mid (objprocess.creationdate, 9, 2) & ":" & _
 	            Mid(objprocess.creationdate, 11, 2) & ":" & Mid(objprocess.creationdate,13, 2))

	if len (strNameOfUser) > 1 then
		LoggedInCounter = LoggedInCounter + 1
		UserListing = UserListing & "<tr><td>" &  Servername & "</td><td>" & strOwner & "</td><td>" & Startdate & "</td></tr>"
	end if

	Next
else
	BadWMI = BadWMI & ServerName & "<br>"
	
End If

on error goto 0
End Function


Function CheckStatus(strAddress)
	Dim objPing, objRetStatus
	Set objPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
      ("select * from Win32_PingStatus where address = '" & strAddress & "'")
	For Each objRetStatus In objPing
	        If IsNull(objRetStatus.StatusCode) Or objRetStatus.StatusCode <> 0 Then
				CheckStatus = False
				DeadServerCount = DeadServerCount + 1
	        Else
				CheckStatus = True
	        End If
	Next
	Set objPing = Nothing
End Function


Function SendMail(strRecipient, strHeader)
	Set objMessage = CreateObject("CDO.Message")
	objMessage.Subject = strHeader
	objMessage.From = "RDPSessionScanner@MyEmailDomain.COM"
	objMessage.To = strRecipient
	'objMessage.HTMLBody = replace(mailmessage,vbCrLf,"<br>")

	objMessage.HTMLBody = MailMessage
	
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2

	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MySMTPServer

	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25

	objMessage.Configuration.Fields.Update
	objMessage.Send
	Set objMessage = Nothing
End Function
