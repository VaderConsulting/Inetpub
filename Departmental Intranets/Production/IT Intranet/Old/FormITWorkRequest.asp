<%@Language=VBScript%>
<!-- #Include file="Includes/Menu.asp" -->
<!-- #Include file="Includes/FormatDate.asp" -->
<%
	Response.Expires = 0
	
	'check if the user has clicked the submit button
	If Len(Request.Form("SubmitDate")) > 0 Then

		Dim lstrHTML
		lstrHTML = ""
		lstrHTML = lstrHTML & "<HTML>" & vbcrlf
		lstrHTML = lstrHTML & "<HEAD>" & vbcrlf
		lstrHTML = lstrHTML & "<STYLE>" & vbcrlf
		lstrHTML = lstrHTML & "BODY, TABLE" & vbcrlf
		lstrHTML = lstrHTML & "{" & vbcrlf
		lstrHTML = lstrHTML & "	COLOR:				#555555;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-FAMILY:		Arial, Verdana, Helvetica;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-SIZE:			8pt;" & vbcrlf
		lstrHTML = lstrHTML & "}" & vbcrlf
		lstrHTML = lstrHTML & ".MainHeading" & vbcrlf
		lstrHTML = lstrHTML & "{" & vbcrlf
		lstrHTML = lstrHTML & "	COLOR:				#3C68AE;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-SIZE:			12pt;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-WEIGHT:		Bold;" & vbcrlf
		lstrHTML = lstrHTML & "	HEIGHT:				23px;" & vbcrlf

		lstrHTML = lstrHTML & "}" & vbcrlf
		lstrHTML = lstrHTML & ".SubHeading" & vbcrlf
		lstrHTML = lstrHTML & "{" & vbcrlf
		lstrHTML = lstrHTML & "	COLOR:				#3C68AE;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-SIZE:			11px;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-WEIGHT:		Bold;" & vbcrlf
		lstrHTML = lstrHTML & "	TEXT-TRANSFORM:		uppercase;" & vbcrlf
		lstrHTML = lstrHTML & "}" & vbcrlf
		lstrHTML = lstrHTML & "</STYLE>" & vbcrlf
		lstrHTML = lstrHTML & "</HEAD>" & vbcrlf
		lstrHTML = lstrHTML & "<BODY>" & vbcrlf
		lstrHTML = lstrHTML & "	<TABLE cellpadding=0 cellspacing=0 width=570>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2><span style=""color:black;"">IT</span>&nbsp;work request form</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		'now build the HTML to send to the help-desk
		lstrHTML = lstrHTML & "		<TR><TD><B>Submit Date:</B></TD><TD>" & Request.Form("SubmitDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Required By Date:</B></TD><TD>" & Request.Form("RequiredDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>User Account:</B></TD><TD>" & Request.Form("UserName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Requestor:</B></TD><TD>" & Request.Form("Requestor") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD colspan=2><B>Brief Description:</B></TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD colspan=2>" & request.form("Title") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD colspan=2><B>Detailed Description of Requirements:</B></TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD colspan=2>"
		
		'check if a reason was entered
		If Len(Request.Form("Reason")) = 0 Then
			lstrHTML = lstrHTML & "No reason specified"
		Else
			lstrHTML = lstrHTML & Request.Form("Reason")
		End If
		
		lstrHTML = lstrHTML & "			</TD></TR>"
		lstrHTML = lstrHTML & "		</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'now write the file
		Dim OutputPath
		OutputPath = Server.MapPath(Application("Documents") & "Forms/ITWorkRequest")
		Dim FileName
		FileName = "\" & FormatDate(Now, C_FormatDate_Identifier) & " " & Request.Form("Title") & ".html"
		Dim FS, File
		Set FS = Server.CreateObject("Scripting.FileSystemObject")
		Set File = FS.CreateTextFile(OutputPath & FileName)
		File.write lstrHTML
		File.Close
		Set File = Nothing
		Set FS = Nothing
		
		'Build the email
		Dim lstrFrom
		Dim lstrTo
		Dim lstrSubject

		'to & subject
		lstrFrom = Application("EmailFromAccount")
		lstrTo = Application("EmailWorkRequest")
		lstrSubject = EmailSubjectDev & "IT Work Request"
		lstrServer =  Application("EmailSMTPServer")

		'Email constants
		Const CdoBodyFormatHTML	= 0
		
		On Error Resume Next
		
		'create the CDO object
		Dim Mail
		Set Mail = Server.CreateObject("CDO.Message")
		
		If Err.number <> 0 Then
			Response.Write "Create CDO.Message Object: " & Err.number & " - " & Err.Description
			Response.End
		End If
		
		'assign the values to the email object
		Mail.To = lstrTo
		Mail.From = lstrFrom
		'Mail.CC = Session("EmailAddr")
		Mail.Subject = lstrSubject
		Mail.HTMLBody = lstrHTML
		
		Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = lstrServer
		Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		Mail.Configuration.Fields.Update()
		
		If Err.number <> 0 Then
			Response.Write "Building the email content: " & Err.number & " - " & Err.Description
			Response.End
		End If
		
		'send the email
		Mail.Send
		
		If Err.number <> 0 Then
			Response.Write "Sending the email: " & Err.number & " - " & Err.Description
			Response.End
		End If
		
		'release the CDO object
		set mail = nothing
		On Error Goto 0
		
		If Err.number <> 0 Then
			Response.Write "Other: " & Err.number & " - " & Err.Description
			Response.End
		End If
%>
<HTML>
<HEAD>
	<TITLE>IT Intranet: IT Work Request Form</TITLE>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
</HEAD>
<BODY>
	<TABLE width=770 cellpadding=0 cellspacing=0 align=center width=770>
		<TR id=HidePrint height=79><TD colspan=3 background=Images/Logo.gif align=right class=Copyright><A href="SiteMap.asp">Site Map</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="ContactUs.asp">Contact Us</A></TD></TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint>
			<TD width=190><DIV class=HomeLink><A href="index.asp" width=100%>HOME</A></DIV></TD>
			<TD width=10>&nbsp;</TD>
			<TD width=570>&nbsp;</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR>
			<TD id=HidePrint width=190 valign=top>
				<TABLE cellpadding=0 cellspacing=0 width=100%>
					<%ShowMenu%>
				</TABLE>
			</TD>
			<TD width=10></TD>
			<TD width=570 valign=top>
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">IT Work Request Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>IT Work Request Submitted</P>
				<P>Your request has been emailed to the IT Helpdesk management for review.</P>
				<DIV class=Copyright><BR>This page last updated: 2 Aug 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>		
<%		
	Else
		
		Dim ServiceAccount
		ServiceAccount = Application("ADServiceAccount")
		Dim ServicePassword
		ServicePassword = Application("ADServicePassword")
		Const ADS_SECURE_AUTHENTICATION = 1
		Const ADS_USE_ENCRYPTION = 2
		Const ADS_PROPERTY_CLEAR = 1
	
		Dim lstrEmail
		Dim lstrUserFullName
		Dim lstrDisplayUser
		lstrDisplayUser = Request.ServerVariables("LOGON_USER")
		lstrDisplayUser = UCASE(Replace(lstrDisplayUser, "OXIANA\", ""))
	
		'define the recordset objects
		Dim adoConn
		Dim adoCmd
		Dim adoRS
	
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
		adoCmd.CommandText = "<LDAP://DC=Oxiana,DC=com,DC=au>;(&(objectCategory=user)(sAMAccountName=" + lstrDisplayUser + "));ADsPath,displayName,CN,givenName,mail,sn,telephoneNumber,company,st,streetAddress,facsimileTelephoneNumber,description,postalcode,physicalDeliveryOfficeName,mobile,l,title,Department,countryCode;subtree"
		On Error Resume Next
		Set adoRS = adoCmd.Execute
		On Error Goto 0
	
		If UCase(typename(adoRS)) = "EMPTY" Then
			lstrUserFullName = "Unknown User"
			lstrEmail = ""
		Elseif adoRS.EOF Then
			lstrUserFullName = "Unknown User"
			lstrEmail = ""
		else
			lstrUserFullName = adoRS.fields("displayName")
			lstrEmail = adoRS.fields("mail")
		End If
		
		'close the recordset objects
		if UCase(TypeName(adoRS)) <> "EMPTY" Then
			adoRS.Close
		End If
		adoConn.Close
		
		'release the objects
		Set adoRS = Nothing
		Set adoConn = Nothing
%>
<HTML>
<HEAD>
	<TITLE>IT Intranet: IT Work Request Form</TITLE>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
	<SCRIPT language=JavaScript>
	<!--
	function HideNonPrintables()
	{
		for (Counter=0;Counter<HidePrint.length;Counter++) HidePrint[Counter].style.display = 'None';
	}
	function ShowNonPrintables()
	{
		for (Counter=0;Counter<HidePrint.length;Counter++) HidePrint[Counter].style.display = 'Inline';
	}
	//-->
	</SCRIPT>
	<SCRIPT language=javascript src="Scripts/HTMLDateFormat.js"></SCRIPT>
</HEAD>

<BODY onbeforeprint="HideNonPrintables()" onafterprint="ShowNonPrintables()">
	<TABLE width=770 cellpadding=0 cellspacing=0 align=center width=770>
		<TR id=HidePrint height=79><TD colspan=3 background=Images/Logo.gif align=right class=Copyright><A href="SiteMap.asp">Site Map</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="ContactUs.asp">Contact Us</A></TD></TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint>
			<TD width=190><DIV class=HomeLink><A href="index.asp" width=100%>HOME</A></DIV></TD>
			<TD width=10>&nbsp;</TD>
			<TD width=570>&nbsp;</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR>
			<TD id=HidePrint width=190 valign=top>
				<TABLE cellpadding=0 cellspacing=0 width=100%>
					<%ShowMenu%>
				</TABLE>
			</TD>
			<TD width=10></TD>
			<TD width=570 valign=top>
				<FORM method=post action="FormITWorkRequest.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">IT Work Request Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
							</TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=Warning>Requests for work will normally be subject to a 7-day turnaround depending on the nature of the request. If your request is urgent, please make a note of this in the "Detailed Description of Requirement" field.</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Submit Date:</B></TD>
							<TD><INPUT name=SubmitDate value="<%= FormatDate(Now, C_FormatDate_ShortDate)%>" onblur="CheckDate(this, true)" style="WIDTH:75px" readonly></TD>
						</TR>
						<TR>
							<TD><B>Required By Date:</B></TD>
							<TD><INPUT name=RequiredDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR>
							<TD><B>Requestor:</B></TD>
							<TD><INPUT name=Requestor value="<%= lstrUserFullName%>" readonly><INPUT type=hidden name=UserName value="<%= Request.ServerVariables("LOGON_USER")%>"><INPUT type=hidden name=Email value="<%= lstrEmail%>"></TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Brief Description:</B></TD></TR>
						<TR>
							<TD colspan=2><INPUT name=Title style="WIDTH:570px"></TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Detailed Description Of Requirement:</B></TD></TR>
						<TR>
							<TD colspan=2><TEXTAREA name=Reason style="HEIGHT:100px;WIDTH:570px"></TEXTAREA></TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 align=center><INPUT type=button value=Submit onclick="form.submit()"></TD></TR>
					</TABLE>
				</FORM>
				<DIV class=Copyright><BR>This page last updated: 10 Sep 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>
<%
	End If
%>