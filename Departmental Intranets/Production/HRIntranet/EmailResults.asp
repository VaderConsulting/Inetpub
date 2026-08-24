<%@Language = VBScript%>

<%
Dim ServiceAccount
ServiceAccount = Application("ADServiceAccount")
Dim ServicePassword
ServicePassword = Application("ADServicePassword")

Sub ShowResults()
	
	Dim FirstName
	Dim LastName
	Dim RecordMatch
	Dim lngUAC
	Dim Matches()

	FirstName = UCase(Request.Form("FirstName"))
	LastName = UCase(Request.Form("LastName"))
	
	If Len(FirstName) = 0 And Len(LastName) = 0 Then
		response.write "<TR><TD>Please specify a first or last name to search for.</TD></TR>"
	Else

		'create the connection object
		Set adoConn = CreateObject("ADODB.Connection")
		
		'Attempt to locate the user account in AD
		adoConn.Provider = "ADsDSOObject"
		adoConn.Properties("User ID") = ServiceAccount
		adoConn.Properties("Password") = ServicePassword
		adoConn.Properties("ADSI Flag") = 1
		adoConn.Properties("Encrypt Password") = True
		adoConn.Open "ADSI"
		
		Set adoCmd = CreateObject("ADODB.Command")
		Set adoCmd.ActiveConnection = adoConn
		adoCmd.Properties("Cache results") = True
		adoCmd.Properties("Page Size") = 1000
		adoCmd.CommandText = "<LDAP://OU=User Accounts,DC=Oxiana,DC=com,DC=au>;(&(objectCategory=user)(objectClass=user));userAccountControl,canonicalName,distinguishedname,givenname,sn,title,mail;subtree"
		Set adoRS = adoCmd.Execute
		
		Redim Matches(2,0)
		While Not adoRS.Eof
			
			lngUAC = CLng(adoRS("userAccountControl"))
  			' Sanity Check
			' 1. Surname must be longer than 0 chars
			' 2. Given Name must be longer than 0 chars
			' 3. Account must not be disabled
			' 4. mail cannot be blank
			
			If (Len(adoRS.Fields("SN")) > 0) And (Len(adoRS.Fields("GivenName")) > 0) And ((lngUAC And 2) = 0) And (Len(Trim(adoRS("Mail"))) > 0) Then
				
					
				
				RecordMatch = True
				
				'check if the user has specified both 
				If (Len(FirstName) > 0) Then
					If UCase(Left(adoRS("givenName"), Len(FirstName))) <> FirstName Then
						RecordMatch = False
					End If
				End If
				
				If (Len(LastName) > 0) Then
					If UCase(Left(adoRS("sn"), Len(LastName))) <> LastName Then
						RecordMatch = False
					End If
				End If
				
				If RecordMatch = True Then

					Redim Preserve Matches(2, UBound(Matches, 2) + 1)

					Matches(0, UBound(Matches, 2)) = UCase(adoRS("sn") & ", " & adoRS("givenName"))
					Matches(1, UBound(Matches, 2)) = adoRS("sn") & ", " & adoRS("givenName")
					Matches(2, UBound(Matches, 2)) = Replace(adoRS("Mail"), "'", "\'")
				End If
				
			End If
			
			adoRS.MoveNext
		Wend
		adoRS.Close

Dim strMisc
Dim intCount
Dim intCount2
Dim intCount3

' Bubble Sort it....Sure it takes a while but it's quick enough at the moment.
  For intCount = 0 To UBound(Matches,2)
    For intCount2 = intCount To UBound(Matches,2)
      If Matches(0,intCount) > Matches(0,intCount2) Then
        For intCount3 = 0 To 2
          strMisc = Matches(intCount3,intCount)
          Matches(intCount3,intCount) = Matches(intCount3,intCount2)
          Matches(intCount3,intCount2) = strMisc
        Next
      End If
    Next
  Next
		RecordMatch = False
		For Counter = 1 to UBound(Matches, 2)
			Response.write "<TR><TD class=""Copyright"" style=""Text-align:left""><A href=""#"" onclick=""return AddEmail('" & Matches(2, Counter) & "')"">" & Matches(1, Counter) & "</A></TD><TD>" & Matches(2, Counter) & "</TD></TR>" & vbcrlf
			RecordMatch = True
		Next
		If RecordMatch = False Then
			Response.write "<TR><TD>No matching usernames were found.</TD></TR>"
		End If
	End If
End Sub
%>
<html>
<head>
	<LINK rel="Stylesheet" type="text/css" href="Styles/Style.css"></LINK>
	<script language="JavaScript">
	function AddEmail(Address)
	{
		window.parent.AddEmail(Address);
		return false;
	}
	</script>
</head>

<body bottommargin="0" leftmargin="0" rightmargin="0" topmargin="0">
	<TABLE cellpadding="3" cellspacing="1">
		<%ShowResults()%>
	</TABLE>
</body>
</html>