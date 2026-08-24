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
		lstrSubject = EmailSubjectDev & "Change Employee Personal Details"
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
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2>Change Employee Personal Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		'EMPLOYEE DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Employee Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Last Name:</B></TD><TD>" & Request.Form("LastName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>First Name:</B></TD><TD>" & Request.Form("FirstName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Middle Name:</B></TD><TD>" & Request.Form("MiddleName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>ID Number:</B></TD><TD>" & Request.Form("IDNumber") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Department:</B></TD><TD>" & Request.Form("Department") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>New Address:</B></TD><TD>" & Request.Form("NewAddress") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>City:</B></TD><TD>" & Request.Form("City") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Post Code:</B></TD><TD>" & Request.Form("PostCode") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Home Phone:</B></TD><TD>" & Request.Form("HomePhone") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Is this a Silent Number:</B></TD><TD>" & Request.Form("SilentNumber") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Mobile Phone:</B></TD><TD>" & Request.Form("MobilePhone") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Change Date:</B></TD><TD>" & Request.Form("Date") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'BANK DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Bank Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD><TD><B><SPAN style=""WIDTH:80px"">BSB</SPAN>Account No</B></TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Account 1:</B></TD><TD><SPAN style=""WIDTH:80px"">" & Request.Form("BSB1") & "</SPAN>" & Request.form("AccountNo1") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Account 2:</B></TD><TD><SPAN style=""WIDTH:80px"">" & Request.Form("BSB2") & "</SPAN>" & Request.form("AccountNo2") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Account 3:</B></TD><TD><SPAN style=""WIDTH:80px"">" & Request.Form("BSB3") & "</SPAN>" & Request.form("AccountNo3") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Account 4:</B></TD><TD><SPAN style=""WIDTH:80px"">" & Request.Form("BSB4") & "</SPAN>" & Request.form("AccountNo4") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Account 5:</B></TD><TD><SPAN style=""WIDTH:80px"">" & Request.Form("BSB5") & "</SPAN>" & Request.form("AccountNo5") & "</TD></TR>" & vbcrlf

		lstrHTML = lstrHTML & "	</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'save the file
		SaveToFile Server.MapPath(Application("Documents") & "Forms\ChangeOfAddress") & "\" & FormatDate(Now, C_FormatDate_Identifier) & " " & replace(Request.Form("Requestor"), "\", "_") & ".html", lstrHTML
		
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
	<TITLE>HR Intranet: Change Employee Personal Details</TITLE>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Change Employee Personal Details</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>Change of Address Request Submitted</P>
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
	<TITLE>HR Intranet: Change Employee Personal Details</TITLE>
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
		
		//if (CheckField(TheForm.AuthorisedBy, '', 'You must specify the person who has authorised this request') == false) return false;
		//if (CheckField(TheForm.Position, '', 'Please specify the position of the person who authorised this request') == false) return false;
		//if (CheckField(TheForm.RequiredDate, '', 'Please specify the date that this request must be completed by') == false) return false;
		//if (CheckField(TheForm.AuthTelephone, '', 'Please specify a telephone number for the person who authorised this request') == false) return false;
		//if (CheckField(TheForm.Site, '', 'Please specify the site location of the person who authorised this request') == false) return false;
		
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
				<FORM id=Form1 method=post action="FormChangeOfAddress.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Change Employee Personal Details</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
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
							<TD><B>ID Number:</B></TD>
							<TD><INPUT name=IDNumber></TD>
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
							<TD><B>New Address:</B></TD>
							<TD><INPUT name=NewAddress></TD>
						</TR>
						<TR>
							<TD><B>City:</B></TD>
							<TD><INPUT name=City></TD>
						</TR>
						<TR>
							<TD><B>Post Code:</B></TD>
							<TD><INPUT name=PostCode></TD>
						</TR>
						<TR>
							<TD><B>Home Phone:</B></TD>
							<TD><INPUT name=HomePhone></TD>
						</TR>
						<TR>
							<TD><B>Is this a Silent Number?</B></TD>
							<TD>
								<SELECT name=SilentNumber>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR>
							<TD><B>Mobile Phone:</B></TD>
							<TD><INPUT name=MobilePhone></TD>
						</TR>
						<TR>
							<TD><B>Change Date:</B></TD>
							<TD><INPUT name=Date onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>

<!-- EMPLOYEE DETAILS-->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Bank Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD><TD><B><SPAN style="WIDTH:80px">BSB</SPAN>Account No</B></TD></TR>
						<TR>
							<TD><B>Account 1:</B></TD>
							<TD><INPUT name=BSB1 style="WIDTH:80px" maxlength=6><INPUT name=AccountNo1 style="WIDTH:150px" maxlength=10></TD>
						</TR>
						<TR>
							<TD><B>Account 2:</B></TD>
							<TD><INPUT name=BSB2 style="WIDTH:80px" maxlength=6><INPUT name=AccountNo2 style="WIDTH:150px" maxlength=10></TD>
						</TR>
						<TR>
							<TD><B>Account 3:</B></TD>
							<TD><INPUT name=BSB3 style="WIDTH:80px" maxlength=6><INPUT name=AccountNo3 style="WIDTH:150px" maxlength=10></TD>
						</TR>
						<TR>
							<TD><B>Account 4:</B></TD>
							<TD><INPUT name=BSB4 style="WIDTH:80px" maxlength=6><INPUT name=AccountNo4 style="WIDTH:150px" maxlength=10></TD>
						</TR>
						<TR>
							<TD><B>Account 5:</B></TD>
							<TD><INPUT name=BSB5 style="WIDTH:80px" maxlength=6><INPUT name=AccountNo5 style="WIDTH:150px" maxlength=10></TD>
						</TR>

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