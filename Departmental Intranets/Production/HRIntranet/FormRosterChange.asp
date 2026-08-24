<%@Language=VBScript%>
<!-- #Include file="Includes/Menu.asp" -->
<!-- #Include file="Includes/FormatDate.asp" -->
<%
	response.expires = 0
'this subroutine will save the passed in string to the specified file
'
'INPUT:
'	pstrFileName	- string containing the filename and path to save the HTML to
'	pstrHTML		- String containing the HTML that is to be saved
Sub SaveToFile(pstrFilename, pstrHTML)
	
	'define the local variables
	Dim FS, File
	
	'create the Filesystemobjects
	Set FS = Server.CreateObject("Scripting.FileSystemObject")
	Set File = FS.CreateTextFile(OutputPath & pstrFileName)
	
	'write the HTML
	File.write pstrHTML
	File.Close
	
	'release the object
	Set File = Nothing
	Set FS = Nothing
End Sub

	Response.Expires = 0
	
	'check if the user has clicked the submit button
	If Len(Request.Form("LastName")) > 0 Then
		
		'Build the HTML body of the email
		Dim lstrFrom
		Dim lstrTo
		Dim lstrCC
		Dim lstrSubject
		Dim lstrHTML
		lstrHTML = ""
		
		'to cc subject
		lstrFrom = Application("EmailFromAccount")	
		lstrTo = Application("EmailITSupport")	
		lstrSubject = EmailSubjectDev & "Roster Change Request"
		lstrServer =  Application("EmailSMTPServer")
		
		lstrCC = Session("EmailAddr")
		if Len(Trim(Request.form("EmailCC"))) > 0 Then
			lstrCC = Trim(Request.Form("EmailCC"))
		End If
		
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
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2>Roster Change Request</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		'now build the HTML to send to the help-desk
		
		'EMPLOYEE DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Employee Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Surname:</B></TD><TD>" & Request.Form("LastName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Given Names:</B></TD><TD>" & Request.Form("FirstName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Site:</B></TD><TD>" & Request.Form("Site") & "</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Roster Transfer</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Current Roster:</B></TD><TD>" & Request.Form("CurrentRoster") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>New Roster:</B></TD><TD>" & Request.Form("NewRoster") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Employee will work same roster as:</B></TD><TD>" & Request.Form("SameRosterAs") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Start/Effective Date:</B></TD><TD>" & Request.Form("StartDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Fly In Date:</B></TD><TD>" & Request.Form("FlyInDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Day in Roster Cycle:</B></TD><TD>" & Request.Form("DayInCycle") & "</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Other Details:</B></TD><TD>" & Replace(Request.Form("OtherDetails"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "	</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'save the file
		SaveToFile Server.MapPath(Application("Documents") & "Forms\RosterChange") & "\" & FormatDate(Now, C_FormatDate_Identifier) & ".html", lstrHTML
		
		'Response.Write lstrHTML
		'Response.End
		
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
		Mail.CC = lstrCC
		Mail.From = lstrFrom
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
	<TITLE>HR Intranet: Roster Change Request</TITLE>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Roster Change Request</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>Roster Change Request Submitted</P>
				<P>Your request has been emailed to the IT Helpdesk. You will be contacted by one of the Helpdesk staff as soon as possible.</P>
				<DIV class=Copyright><BR>This page last updated: 22 Dec 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>		
<%		
	Else

		lstrUserFullName = Session("SAMUserName")
		lstrEmail = Session("EmailAddr")
%>
<HTML>
<HEAD>
	<TITLE>HR Intranet: Roster Change Request</TITLE>
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
	<SCRIPT language=JavaScript>
	<!--
	function ValidateAndSubmit()
	{
		var TheForm = document.forms.Form1;
		
		if (CheckField(TheForm.LastName, '', 'Please specify the users Surname') == false) return false;
		if (CheckField(TheForm.FirstName, '', 'Please specify the users Given Names') == false) return false;
		if (CheckField(TheForm.Site, '', 'Please specify the users Site') == false) return false;
		
		TheForm.submit();
	}
	function CheckField(TheField, BadValue, Message)
	{
		if (TheField.value == BadValue)
		{
			alert(Message);
			TheField.focus();
			return false;
		}
	}
	function PromptForEmail()
	{
		var TheForm = document.forms.Form1;
		var Temp = window.showModalDialog('EmailValidator.asp', TheForm.EmailCC.value, 'dialogheight:350px;dialogwidth:400px;status:no');
		if (Temp != null)
		{
			TheForm.EmailCC.value = Temp;
		}
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
				<FORM id=Form1 method=post action="FormRosterChange.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Roster Change Request</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
							</TD>
						</TR>

<!-- EMPLOYEE DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Employee Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Surname: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=LastName></TD>
						</TR>
						<TR>
							<TD><B>Given Names: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=FirstName></TD>
						</TR>
						<TR>
							<TD><B>Site/Location: <SPAN class="Warning">*</SPAN></B></TD>
							<TD>
								<SELECT name=Site>
									<OPTION value="">-- Select a Site --</OPTION>
									<OPTION value="Golden Grove">Golden Grove</OPTION>
									<OPTION value="Melbourne">Melbourne</OPTION>
									<OPTION value="Adelaide">Adelaide</OPTION>
									<OPTION value="SPDG">SPDG</OPTION>
									<OPTION value="Other">Other...</OPTION>
								</SELECT>
							</TD>
						</TR>

<!-- ROSTER TRANSFER -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Roster Transfer</TD></TR>
						<TR>
							<TD><B>Current Roster:</B></TD>
							<TD>
								<INPUT name=CurrentRoster>
							</TD>
						</TR>
						<TR>
							<TD><B>New Roster:</B></TD>
							<TD>
								<INPUT name=NewRoster>
							</TD>
						</TR>
						<TR>
							<TD><B>Employee will work same roster as:</B></TD>
							<TD>
								<INPUT name=SameRosterAs>
							</TD>
						</TR>
						<TR>
							<TD><B>Start/Effective Date:</B></TD>
							<TD>
								<INPUT name=StartDate onblur="CheckDate(this, true)" style="WIDTH:75px">
							</TD>
						</TR>
						<TR>
							<TD><B>Fly In Date:</B></TD>
							<TD>
								<INPUT name=FlyInDate onblur="CheckDate(this, true)" style="WIDTH:75px">
							</TD>
						</TR>
						<TR>
							<TD><B>Day in Roster Cycle:</B></TD>
							<TD>
								<INPUT name=DayInCycle>
							</TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Other Details:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=OtherDetails style="WIDTH:450px;HEIGHT:75px"></TEXTAREA></TD></TR>

<!-- EMAILCC -->
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Email CC:</B></TD>
							<TD><INPUT name=EmailCC readonly style="BACKGROUND-COLOR:E8E8E8;width:200px"><INPUT type=button value="..." onclick="PromptForEmail()"></TD>
						</TR>
						<TR>
							<TD colspan=2>(If you wish to send this form to an additional email address, enter the email address here. Please ensure you enter the correct address.)</TD>
						</TR>
<!-- SUBMIT -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD align=center colspan=2><INPUT type=button value=Submit onclick="ValidateAndSubmit()"></TD></TR>
					</TABLE>
				</FORM>
				<DIV class=Copyright><BR>This page last updated: 22 Dec 2006 (Empired)</DIV>
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