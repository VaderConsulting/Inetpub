<%@Language=VBScript%>
<!-- #Include file="Includes/Menu.asp" -->
<!-- #Include file="Includes/FormatDate.asp" -->
<%
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

Function ShowExtra(pstrTitle, pstrField, pstrExtra)
	
	Dim lstrTemp
	
	lstrTemp = ""
	lstrTemp = lstrTemp & "		<TR><TD><B>" & pstrTitle & "</B></TD><TD>" & Request.Form(pstrField) 
	If Request.Form(pstrField) = pstrExtra Then
		lstrTemp = lstrTemp & " " & Request.Form(pstrField & "EXTRA") & "</TD></TR>" & vbcrlf
	Else
		lstrTemp = lstrTemp & "</TD></TR>" & vbcrlf
	End If
	
	ShowExtra = lstrTemp
	
End Function

	Response.Expires = 0
	
	'check if the user has clicked the submit button
	If Len(Request.Form("Requestor")) > 0 Then
		
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
		lstrSubject = EmailSubjectDev & "Terminate Employee"
		lstrServer =  Application("EmailSMTPServer")
		
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
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2>Terminate Employee</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		'REQUEST AUTHORISATION DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Request Authorisation Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Authorised By:</B></TD><TD>" & Request.Form("AuthorisedBy") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Authorised By Account:</B></TD><TD>" & Session("SAMUserName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Position:</B></TD><TD>" & Request.Form("Position") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Required Date:</B></TD><TD>" & Request.Form("RequiredDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Telephone:</B></TD><TD>" & Request.Form("AuthTelephone") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Site:</B></TD><TD>" & Request.Form("Site") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'USER DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>User Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Last Name:</B></TD><TD>" & Request.Form("LastName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>First Name:</B></TD><TD>" & Request.Form("FirstName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Middle Name:</B></TD><TD>" & Request.Form("MiddleName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Date Of Birth:</B></TD><TD>" & Request.Form("DOB") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Telephone:</B></TD><TD>" & Request.Form("AuthTelephone") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Department:</B></TD><TD>" & Request.Form("Department") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Position Title:</B></TD><TD>" & Request.Form("PositionTitle") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Position No:</B></TD><TD>" & Request.Form("PositionNo") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Position to be refilled?</B></TD><TD>" & Request.Form("PositionRefilled") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Ellipse User Name:</B></TD><TD>" & Request.Form("EllipseUserName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Termination Type:</B></TD><TD>" & Request.Form("TerminationType") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Last Day of Employment:</B></TD><TD>" & Request.Form("TerminationDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'NETWORK DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Network Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>User Account:</B></TD><TD>" & Request.Form("DisableUser")
		If Request.Form("DisableUser") = "Disable on Date..." Then
			lstrHTML = lstrHTML & " " & Request.Form("DisableUserEXTRA") & "</TD></TR>" & vbcrlf
		Else
			lstrHTML = lstrHTML & "</TD></TR>" & vbcrlf
		End If
		
		lstrHTML = lstrHTML & "		<TR><TD><B>Users Home Drive:</B></TD><TD>" & Request.Form("UserHomeDrive") 
		If Request.Form("UserHomeDrive") = "Move to new location..." Then
			lstrHTML = lstrHTML & " " & Request.Form("UserHomeDriveEXTRA") & "</TD></TR>" & vbcrlf
		Else
			lstrHTML = lstrHTML & "</TD></TR>" & vbcrlf
		End If
		
		lstrHTML = lstrHTML & "		<TR><TD><B>Users Mail Box:</B></TD><TD>" & Request.Form("UserMailbox") 
		If Request.Form("UserMailbox") = "Copy to CD & forward to Manager" Then
			lstrHTML = lstrHTML & " - " & Request.Form("UserMailboxEXTRA") & "</TD></TR>" & vbcrlf
		Else
			lstrHTML = lstrHTML & "</TD></TR>" & vbcrlf
		End If
		
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'APPLICATION DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Application Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & ShowExtra("Ellipse:", "Ellipse", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("Hyperion Enterprise:", "HE", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("ANZ Online:", "ANZOnline", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("Cognos ePlanning:", "Cognos", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("DirectLink (NBNZ Online - Adelaide):", "DirectLinkAdelaide", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("DirectLink (NBNZ Online - Waihi):", "DirectLinkWaihi", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("Hyperion Financial (AHFM):", "AHFM", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("Hyperion Planning:", "HP", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("ESS (Contractor Only):", "ESS", "Disable on Date...")
		lstrHTML = lstrHTML & ShowExtra("SAP Read Only Access:", "SAP", "Disable on Date...")
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'ADDITIONAL INFORMATION
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Additional Information</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Other Details:</B></TD><TD>" & Replace(Request.Form("AdditionalText"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "	</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'save the file
		SaveToFile Server.MapPath(Application("Documents") & "Forms\TerminatingEmployee") & "\" & FormatDate(Now, C_FormatDate_Identifier) & " " & replace(Request.Form("Requestor"), "\", "_") & ".html", lstrHTML
		
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
	<TITLE>HR Intranet: Terminate Employee Form</TITLE>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
</HEAD>
<BODY>
	<TABLE width=770 cellpadding=0 cellspacing=0 align=center width=770>
		<TR id=HidePrint height=79><TD colspan=3 background=Images/Logo.gif align=right class=Copyright><A href="SiteMap.asp">Site Map</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="ContactUs.asp">Contact Us</A></TD></TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint>
			<TD width=190><DIV class=HomeLink><A onclick="window.close()" width=100%>Back to HR Intranet Home</A></DIV></TD>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Terminate Employee Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>Terminate Employee Access Request Submitted</P>
				<P>Your request has been emailed to the IT Helpdesk. You will be contacted by one of the Helpdesk staff as soon as possible.</P>
				<DIV class=Copyright><BR>This page last updated: 18 Dec 2006 (Empired)</DIV>
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
	<TITLE>HR Intranet: Terminate Employee Form</TITLE>
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
	function ShowEXTRA(TheCombo, ExtraValue)
	{
		if (TheCombo.value == ExtraValue)
		{
			eval('document.forms.Form1.' + TheCombo.name + 'EXTRA.style.display = \'inline\'');
			eval('document.forms.Form1.' + TheCombo.name + 'EXTRA.focus()');
		}
		else
		{
			eval('document.forms.Form1.' + TheCombo.name + 'EXTRA.style.display = \'none\'');
		}
	}
	function ValidateAndSubmit()
	{
		var TheForm = document.forms.Form1;
		
		if (CheckField(TheForm.AuthorisedBy, '', 'You must specify the person who has authorised this request') == false) return false;
		if (CheckField(TheForm.Position, '', 'Please specify the position of the person who authorised this request') == false) return false;
		if (CheckField(TheForm.RequiredDate, '', 'Please specify the date that this request must be completed by') == false) return false;
		if (CheckField(TheForm.AuthTelephone, '', 'Please specify a telephone number for the person who authorised this request') == false) return false;
		if (CheckField(TheForm.Site, '', 'Please specify the site location of the person who authorised this request') == false) return false;
		
		if (CheckField(TheForm.LastName, '', 'Please specify the users Last Name') == false) return false;
		if (CheckField(TheForm.FirstName, '', 'Please specify the users First Name') == false) return false;
		
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
	//-->
	</SCRIPT>
	<SCRIPT language=javascript src="Scripts/HTMLDateFormat.js"></SCRIPT>
</HEAD>

<BODY onbeforeprint="HideNonPrintables()" onafterprint="ShowNonPrintables()">
	<TABLE width=770 cellpadding=0 cellspacing=0 align=center width=770>
		<TR id=HidePrint height=79><TD colspan=3 background=Images/Logo.gif align=right class=Copyright><A href="SiteMap.asp">Site Map</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="ContactUs.asp">Contact Us</A></TD></TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint>
			<TD width=190><DIV class=HomeLink><A onclick="window.close()" width=100%>Back to HR Intranet Home</A></DIV></TD>
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
				<FORM id=Form1 method=post action="FormTerminatingEmployee.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Terminate Employee Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
							</TD>
						</TR>

<!-- REQUEST AUTHORISATION DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Request Authorisation Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD width=200><B>Requestor: <SPAN class="Warning">*</SPAN></B></TD>
							<TD width=370><INPUT name=Requestor readonly style="WIDTH:150px" value="<%= lstrUserFullName%>"></TD>
						</TR>
						<TR>
							<TD width=200><B>Authorised By: <SPAN class="Warning">*</SPAN></B></TD>
							<TD width=370><INPUT name=AuthorisedBy style="WIDTH:150px"></TD>
						</TR>
						<TR>
							<TD><B>Position: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=Position></TD>
						</TR>
						<TR>
							<TD><B>Required Date: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=RequiredDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR>
							<TD><B>Telephone: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=AuthTelephone value=""></TD>
						</TR>
						<TR>
							<TD><B>Site: <SPAN class="Warning">*</SPAN></B></TD>
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

<!-- EMPLOYEE DETAILS-->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Employee Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Last Name: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=LastName></TD>
						</TR>
						<TR>
							<TD><B>First Name: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=FirstName></TD>
						</TR>
						<TR>
							<TD><B>Middle Name:</B></TD>
							<TD><INPUT name=MiddleName></TD>
						</TR>
						<TR>
							<TD><B>Date Of Birth:</B></TD>
							<TD><INPUT name=DOB onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR>
							<TD><B>Department:</B></TD>
							<TD>
								<SELECT name=Department>
									<OPTION value="">-- Select a Department --</OPTION>
									<OPTION value="Melbourne - Finance">Melbourne - Finance</OPTION>
									<OPTION value="Melbourne - Human Resources">Melbourne - Human Resources</OPTION>
									<OPTION value="Melbourne - Company Secretary">Melbourne - Company Secretary</OPTION>
									<OPTION value="Melbourne - Marketing">Melbourne - Marketing</OPTION>
									<OPTION value="Melbourne - Business Development">Melbourne - Business Development</OPTION>
									<OPTION value="Melbourne - Corporate Development">Melbourne - Corporate Development</OPTION>
									<OPTION value="Melbourne - Technical & Projects<">Melbourne - Technical & Projects</OPTION>
									<OPTION value="Melbourne - Executive">Melbourne - Executive</OPTION>
									<OPTION value="Golden Grove - HR/Administration">Golden Grove - HR/Administration</OPTION>
									<OPTION value="Golden Grove - Engineering">Golden Grove - Engineering</OPTION>
									<OPTION value="Golden Grove - Environmental & Community Relations">Golden Grove - Environmental & Community Relations</OPTION>
									<OPTION value="Golden Grove - Geology">Golden Grove - Geology</OPTION>
									<OPTION value="Golden Grove - Metallurgy">Golden Grove - Metallurgy</OPTION>
									<OPTION value="Golden Grove - Mining">Golden Grove - Mining</OPTION>
									<OPTION value="Golden Grove - Commercial">Golden Grove - Commercial</OPTION>
									<OPTION value="Golden Grove - Health & Safety">Golden Grove - Health & Safety</OPTION>
									<OPTION value="Prominent Hill - Human Resources">Prominent Hill - Human Resources</OPTION>
									<OPTION value="Prominent Hill - Project Team">Prominent Hill - Project Team</OPTION>
									<OPTION value="Prominent Hill - Processing">Prominent Hill - Processing</OPTION>
									<OPTION value="Prominent Hill - Engineering">Prominent Hill - Engineering</OPTION>
									<OPTION value="Prominent Hill - Executive">Prominent Hill - Executive</OPTION>
									<OPTION value="Prominent Hill - Geology">Prominent Hill - Geology</OPTION>
									<OPTION value="Prominent Hill - Metallurgy">Prominent Hill - Metallurgy</OPTION>
									<OPTION value="Prominent Hill - Mining">Prominent Hill - Mining</OPTION>
									<OPTION value="Prominent Hill - Sustainability">Prominent Hill - Sustainability</OPTION>
									<OPTION value="Prominent Hill - Health & Safety">Prominent Hill - Health & Safety</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR>
							<TD><B>Position Title:</B></TD>
							<TD><INPUT name=PositionTitle></TD>
						</TR>
						<TR>
							<TD><B>Position No:</B></TD>
							<TD><INPUT name=PositionNo></TD>
						</TR>
						<TR>
							<TD><B>Position to be refilled?</B></TD>
							<TD>
								<SELECT name=PositionRefilled>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR>
							<TD><B>Ellipse User Name:</B></TD>
							<TD><INPUT name=EllipseUserName></TD>
						</TR>
						<TR>
							<TD><B>Termination Type:</B></TD>
							<TD>
								<SELECT name=TerminationType>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Dismissed">Dismissed</OPTION>
									<OPTION value="Resigned">Resigned</OPTION>
									<OPTION value="Redundancy">Redundancy</OPTION>
									<OPTION value="Retired">Retired</OPTION>
									<OPTION value="Expired Contract">Expired Contract</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR>
							<TD><B>Last Day of Employment:</B></TD>
							<TD><INPUT name=TerminationDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						
<!-- NETWORK DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Network Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>User Account:</B></TD>
							<TD>
								<SELECT name=DisableUser onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=DisableUserEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>Users Home Drive:</B></TD>
							<TD>
								<SELECT name=UserHomeDrive onchange="ShowEXTRA(this, 'Move to new location...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Delete Immediately">Delete Immediately</OPTION>
									<OPTION value="Copy to CD & forward to Manager">Copy to CD & forward to Manager</OPTION>
									<OPTION value="Move to new location...">Move to new location...</OPTION>
								</SELECT>
								<INPUT name=UserHomeDriveEXTRA style="DISPLAY:none">
							</TD>
						</TR>
						<TR>
							<TD><B>Users Mail Box:</B></TD>
							<TD>
								<SELECT name=UserMailBox onchange="ShowEXTRA(this, 'Copy to CD & forward to Manager')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="PERMANENTLY Deletes">PERMANENTLY Deleted</OPTION>
									<OPTION value="Copy to CD & forward to Manager">Copy to CD & forward to Manager</OPTION>
								</SELECT>
								<SELECT name=UserMailBoxEXTRA style="DISPLAY:None;WIDTH:160px">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Include Personal Folders">Include Personal Folders</OPTION>
									<OPTION value="Dont Include Personal Folders">Do Not Include Personal Folders</OPTION>
								</SELECT>
							</TD>
						</TR>
						
<!-- APPLICATION DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Application Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Ellipse:</B></TD>
							<TD>
								<SELECT name=Ellipse onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=EllipseEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>Hyperion Enterprise:</B></TD>
							<TD>
								<SELECT name=HE onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=HEEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>ANZ Online:</B></TD>
							<TD>
								<SELECT name=ANZOnline onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=ANZOnlineEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>Cognos ePlanning:</B></TD>
							<TD>
								<SELECT name=Cognos onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=CognosEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>DirectLink (NBNZ Online - Adelaide):</B></TD>
							<TD>
								<SELECT name=DirectLinkAdelaide onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=DirectLinkAdelaideEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>DirectLink (NBNZ Online - Waihi):</B></TD>
							<TD>
								<SELECT name=DirectLinkWaihi onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=DirectLinkWaihiEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>Hyperion Financial (AHFM):</B></TD>
							<TD>
								<SELECT name=AHFM onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=AHFMEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>Hyperion Planning (HP):</B></TD>
							<TD>
								<SELECT name=HP onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=HPEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>ESS (Contractor Only):</B></TD>
							<TD>
								<SELECT name=ESS onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=ESSEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						<TR>
							<TD><B>SAP Read Only Access:</B></TD>
							<TD>
								<SELECT name=SAP onchange="ShowEXTRA(this, 'Disable on Date...')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Disable Immediately">Disable Immediately</OPTION>
									<OPTION value="Disable on Date...">Disable on Date...</OPTION>
								</SELECT>
								<INPUT name=SAPEXTRA onblur="CheckDate(this, true)" style="WIDTH:75px;DISPLAY:None">
							</TD>
						</TR>
						
<!-- ADDITIONAL INFORMATION -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Additional Information</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Specify any other details which could not be recorded above:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=AdditionalText style="WIDTH:450px;HEIGHT:75px"></TEXTAREA></TD></TR>
						

<!-- SUBMIT -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD align=center colspan=2><INPUT type=button value=Submit onclick="ValidateAndSubmit()"></TD></TR>
					</TABLE>
				</FORM>
				<DIV class=Copyright><BR>This page last updated: 18 Dec 2006 (Empired)</DIV>
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