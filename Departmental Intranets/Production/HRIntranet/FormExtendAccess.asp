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
		lstrSubject = EmailSubjectDev & "Extend User Access Form (Existing User)"
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
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2>Extend User Access Form (Existing User)</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		'now build the HTML to send to the help-desk
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Request Authorisation Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Authorised By:</B></TD><TD>" & Request.Form("AuthorisedBy") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Authorised By Account:</B></TD><TD>" & Session("SAMUserName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Position:</B></TD><TD>" & Request.Form("Position") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Required Date:</B></TD><TD>" & Request.Form("RequiredDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Telephone:</B></TD><TD>" & Request.Form("AuthTelephone") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Site:</B></TD><TD>" & Request.Form("Site") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>User Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Last Name:</B></TD><TD>" & Request.Form("LastName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>First Name:</B></TD><TD>" & Request.Form("FirstName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Middle Name:</B></TD><TD>" & Request.Form("MiddleName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Date Of Birth:</B></TD><TD>" & Request.Form("DOB") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Telephone:</B></TD><TD>" & Request.Form("Telephone") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Department:</B></TD><TD>" & Request.Form("Department") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Position Name:</B></TD><TD>" & Request.Form("PositionName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Employment Status:</B></TD><TD>" & Request.Form("EmploymentStatus") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Extension Requirements</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Current Contract End Date:</B></TD><TD>" & Request.Form("ContractEnd") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Revised Contract End Date:</B></TD><TD>" & Request.Form("RevisedContractEnd") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Extend Network Access:</B></TD><TD>" & Request.Form("ExtendNetwork") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Extend Ellipse Access:</B></TD><TD>" & Request.Form("ExtendEllipse") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Extend SAP Read Only Access:</B></TD><TD>" & Request.Form("ExtendSAP") & "</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "	</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'save the file
		SaveToFile Server.MapPath(Application("Documents") & "Forms\ExtendAccess") & "\" & FormatDate(Now, C_FormatDate_Identifier) & " " & replace(Request.Form("Requestor"), "\", "_") & ".html", lstrHTML
		
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
	<TITLE>HR Intranet: Extend User Access Form (Existing User)</TITLE>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Extend User Access Form (Existing User)</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>User Access Request Submitted</P>
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
	<TITLE>HR Intranet: Extend User Access Form (Existing User)</TITLE>
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
		
		if (CheckField(TheForm.AuthorisedBy, '', 'You must specify the person who has authorised this request') == false) return false;
		if (CheckField(TheForm.Position, '', 'Please specify the position of the person who authorised this request') == false) return false;
		if (CheckField(TheForm.RequiredDate, '', 'Please specify the date that this request must be completed by') == false) return false;
		if (CheckField(TheForm.AuthTelephone, '', 'Please specify a telephone number for the person who authorised this request') == false) return false;
		if (CheckField(TheForm.Site, '', 'Please specify the site location of the person who authorised this request') == false) return false;
		
		if (CheckField(TheForm.LastName, '', 'Please specify the users Last Name') == false) return false;
		if (CheckField(TheForm.FirstName, '', 'Please specify the users First Name') == false) return false;
		
		if (CheckField(TheForm.ContractEnd, '', 'Please specify when the contract was due to end') == false) return false;
		if (CheckField(TheForm.RevisedContractEnd, '', 'Please specify the new end date for the contract') == false) return false;
		if (CheckField(TheForm.ExtendNetwork, '', 'Please specify whether Network access needs to be extended') == false) return false;
		if (CheckField(TheForm.ExtendEllipse, '', 'Please specify whether Ellipse access needs to be extended') == false) return false;
		if (CheckField(TheForm.ExtendSAP, '', 'Please specify whether SAP access needs to be extended') == false) return false;
		
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
				<FORM id=Form1 method=post action="FormExtendAccess.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Extend User Access Form (Existing User)</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
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

<!-- USER DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>User Details</TD></TR>
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
							<TD><B>Telephone:</B></TD>
							<TD><INPUT name=Telephone></TD>
						</TR>
						<TR>
							<TD><B>Department:</B></TD>
							<TD>
								<SELECT name=Department>
									<OPTION value="">-- Select a Department --</OPTION>
									<OPTION value="HR/Administration">HR/Administration</OPTION>
									<OPTION value="Engineering">Engineering</OPTION>
									<OPTION value="Environmental & Community Relations">Environmental & Community Relations</OPTION>
									<OPTION value="Geology">Geology</OPTION>
									<OPTION value="Metalurgy">Metalurgy</OPTION>
									<OPTION value="Mining">Mining</OPTION>
									<OPTION value="Commercial">Commercial</OPTION>
									<OPTION value="Health & Safety">Health & Safety</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR>
							<TD><B>Position Name:</B></TD>
							<TD><INPUT name=PositionName></TD>
						</TR>
						<TR>
							<TD><B>Employment Status:</B></TD>
							<TD>
								<SELECT name=EmploymentStatus>
									<OPTION value="">-- Select a Status --</OPTION>
									<OPTION value="Permanent - Full Time">Permanent - Full Time</OPTION>
									<OPTION value="Permanent - Part Time">Permanent - Part Time</OPTION>
									<OPTION value="Fixed Term/Temporary">Fixed Term/Temporary</OPTION>
									<OPTION value="Casual">Casual</OPTION>
									<OPTION value="Contract (Non Payroll)">Contract (Non Payroll)</OPTION>
								</SELECT>
							</TD>
						</TR>

<!-- EXTENSION REQUIREMENTS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Extension Requirements</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Current Contract End Date: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=ContractEnd onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR>
							<TD><B>Revised Contract End Date: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=RevisedContractEnd onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Extend Network Access: <SPAN class="Warning">*</SPAN></B></TD>
							<TD>
								<SELECT name=ExtendNetwork>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR>
							<TD><B>Extend Ellipse Access: <SPAN class="Warning">*</SPAN></B></TD>
							<TD>
								<SELECT name=ExtendEllipse>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR>
							<TD><B>Extend SAP Read Only Access: <SPAN class="Warning">*</SPAN></B></TD>
							<TD>
								<SELECT name=ExtendSAP>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>

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