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
		lstrSubject = EmailSubjectDev & "New Employee Request"
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
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2>New Employee Request</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		'now build the HTML to send to the help-desk
		
		'REQUEST AUTHORISATION DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Request Authorisation Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Authorised By:</B></TD><TD>" & Request.Form("AuthorisedBy") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Authorised By Account:</B></TD><TD>" & Session("SAMUserName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Position:</B></TD><TD>" & Request.Form("Position") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Required Date:</B></TD><TD>" & Request.Form("RequiredDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Telephone:</B></TD><TD>" & Request.Form("AuthTelephone") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Site:</B></TD><TD>" & Request.Form("Site") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'EMPLOYEE DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Employee Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Last Name:</B></TD><TD>" & Request.Form("LastName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Middle Names:</B></TD><TD>" & Request.Form("MiddleNames") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>First Name:</B></TD><TD>" & Request.Form("FirstName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Date of Birth:</B></TD><TD>" & Request.Form("DOB") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'EMPLOYMENT DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Employment Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Employment Status:</B></TD><TD>" & Request.Form("EmploymentStatus") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Department:</B></TD><TD>" & Request.Form("Department") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Direct Supervisor:</B></TD><TD>" & Request.Form("DirectSupervisor") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Cost Centre:</B></TD><TD>" & Request.Form("CostCentre") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Position Title:</B></TD><TD>" & Request.Form("PositionTitle") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Contract Start Date:</B></TD><TD>" & Request.Form("StartDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Contract End Date:</B></TD><TD>" & Request.Form("EndDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Has Prior Network Access?</B></TD><TD>" & Request.Form("PreviousNetwork") & "</TD></TR>" & vbcrlf
		If Request.Form("PreviousNetwork") = "Yes" Then
			lstrHTML = lstrHTML & "		<TR><TD><B>Previous User ID</B></TD><TD>" & Request.Form("PrevUserID") & "</TD></TR>" & vbcrlf
		End If
		lstrHTML = lstrHTML & "		<TR><TD><B>Name of Contractor Company:</B></TD><TD>" & Request.Form("ContractorCompany") & "</TD></TR>" & vbCrLf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'POSITION DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Position Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Existing Position:</B></TD><TD>" & Request.Form("ExistingPosition") & "</TD></TR>" & vbcrlf
		If Request.Form("ExistingPosition") = "Yes" Then
			lstrHTML = lstrHTML & "		<TR><TD><B>Who are they replacing:</B></TD><TD>" & Request.Form("WhoReplacing") & "</TD></TR>" & vbcrlf
		End If
		lstrHTML = lstrHTML & "		<TR><TD><B>Position Number:</B></TD><TD>" & Request.Form("PositionNumber") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Same roster as:</B></TD><TD>" & Request.Form("SameRosterAs") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'NETWORK ACCESS DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Network Access Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Is Network Access Required?</B></TD><TD>" & Request.Form("NetworkAccess") & "</TD></TR>" & vbcrlf
		If Request.Form("NetworkAccess") = "Yes" then
			lstrHTML = lstrHTML & "		<TR><TD><B>Internet Access Required?</B></TD><TD>" & Request.Form("Internet") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>Dial In Access Required?</B></TD><TD>" & Request.Form("DialIn") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>Provide Same Network Access as:</B></TD><TD>" & Request.Form("SameNetworkAccess") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>Specify Folders Required:</B></TD><TD>" & Request.Form("FoldersRequired") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>Provide same email groups as:</B></TD><TD>" & Request.Form("EmailGroups") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>Access to other Mailbox:</B></TD><TD>" & Request.Form("OtherMailbox") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>How will the user connect?</B></TD><TD>" & Request.Form("Connect") & "</TD></TR>" & vbcrlf
		End If
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'ELLIPSE ACCESS DETAILS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Ellipse Access Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Ellipse Access Required?</B></TD><TD>" & Request.Form("Ellipse") & "</TD></TR>" & vbcrlf
		If Request.Form("Ellipse") = "Yes" Then
			lstrHTML = lstrHTML & "		<TR><TD><B>Ellipse Default Printer:</B></TD><TD>" & Request.Form("EllipsePrinter") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>District Access Required:</B></TD><TD>"
			If Len(Request.form("6000")) > 0 Then lstrHTML = lstrHTML & Request.Form("6000") & "<BR>" End If
			If Len(Request.form("6030")) > 0 Then lstrHTML = lstrHTML &  Request.Form("6030") & "<BR>" End If
			If Len(Request.form("6040")) > 0 Then lstrHTML = lstrHTML &  Request.Form("6040") & "<BR>" End If
			If Len(Request.form("6090")) > 0 Then lstrHTML = lstrHTML &  Request.Form("6090") & "<BR>" End If
			If Len(Request.form("6530")) > 0 Then lstrHTML = lstrHTML &  Request.Form("6530") & "<BR>" End If
			If Len(Request.form("6880")) > 0 Then lstrHTML = lstrHTML &  Request.Form("6880") & "<BR>" End If
			If Len(Request.form("6950")) > 0 Then lstrHTML = lstrHTML &  Request.Form("6950") & "<BR>" End If
			If Len(Request.form("ALL")) > 0 Then lstrHTML = lstrHTML &  Request.Form("ALL") & "<BR>" End If
			lstrHTML = lstrHTML & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "<TR><TD><B>Same Access As:</B></TD><TD>" & Request.Form("EllipseSameAccessAs") & "</TD></TR>" & vbcrlf
		End If
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'AUSTRALIAN HYPERION FINANCIAL MANAGEMENT (AHFM)
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Australian Hyperion Financial Management (AHFM)</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>AHFM Access Required?</B></TD><TD>" & Request.Form("AHFM") & "</TD></TR>" & vbcrlf
		If Request.Form("AHFM") = "Yes" Then
			lstrHTML = lstrHTML & "		<TR><TD><B>Provide same access as:</B></TD><TD>" & Request.Form("AHFMAccess") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>Hyperion Reporting Required?</B></TD><TD>" & Request.Form("HyperionReporting") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>Group Access Required for AHFM:</B></TD><TD>"
			If Len(Request.form("AHFMAccess1")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess1") & "<BR>" End If
			If Len(Request.form("AHFMAccess2")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess2") & "<BR>" End If
			If Len(Request.form("AHFMAccess3")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess3") & "<BR>" End If
			If Len(Request.form("AHFMAccess4")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess4") & "<BR>" End If
			If Len(Request.form("AHFMAccess5")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess5") & "<BR>" End If
			If Len(Request.form("AHFMAccess6")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess6") & "<BR>" End If
			If Len(Request.form("AHFMAccess7")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess7") & "<BR>" End If
			If Len(Request.form("AHFMAccess8")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess8") & "<BR>" End If
			If Len(Request.form("AHFMAccess9")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess9") & "<BR>" End If
			If Len(Request.form("AHFMAccess10")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess10") & "<BR>" End If
			If Len(Request.form("AHFMAccess11")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess11") & "<BR>" End If
			If Len(Request.form("AHFMAccess12")) > 0 Then lstrHTML = lstrHTML & Request.Form("AHFMAccess12") & "<BR>" End If
			Response.Write "</TD></TR>" & vbcrlf
		End If
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'OTHER APPLICATIONS
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Other Applications</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>ANZ Online:</B></TD><TD>" & Request.Form("ANZOnline") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Cognos ePlanning:</B></TD><TD>" & Request.Form("CognosePlanning") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>DirectLink (NBNZ Online - Adelaide):</B></TD><TD>" & Request.Form("DirectLinkAdelaide") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>DirectLink (NBNZ Online - Waihi):</B></TD><TD>" & Request.Form("DirectLinkWaihi") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Hyperion Enterprise (HE):</B></TD><TD>" & Request.Form("HE") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Hyperion Planning (HP):</B></TD><TD>" & Request.Form("HP") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>SAP Read Only Access:</B></TD><TD>" & Request.Form("SAP") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'ADDITIONAL INFORMATION
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Additional Information</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Other Details:</B></TD><TD>" & Replace(Request.Form("OtherDetails"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "	</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'save the file
		SaveToFile Server.MapPath(Application("Documents") & "Forms\NewEmployee") & "\" & FormatDate(Now, C_FormatDate_Identifier) & " " & replace(Request.Form("Requestor"), "\", "_") & ".html", lstrHTML
		
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
	<TITLE>HR Intranet: New Employee Request Form</TITLE>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">New Employee Access Request Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>New Employee Request Submitted</P>
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
	<TITLE>HR Intranet: New Employee Request Form</TITLE>
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
	function ShowHide(TheCombo, TheElements)
	{
		var Counter;
		if (TheCombo.value == 'Yes')
		{
			eval('for (Counter=0;Counter<' + TheElements + '.length;Counter++) ' + TheElements + '[Counter].style.display = \'inline\'');
		}
		else
		{
			eval('for (Counter=0;Counter<' + TheElements + '.length;Counter++) ' + TheElements + '[Counter].style.display = \'none\'');
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
		
		if (CheckField(TheForm.NetworkAccess, '', 'Please specify if the user requires Network Access') == false) return false;
		if (CheckField(TheForm.Ellipse, '', 'Please specify if the user requires Ellipse Access') == false) return false;
		if (CheckField(TheForm.AHFM, '', 'Please specify if the user requires AHFM access') == false) return false;
		
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
				<FORM id=Form1 method=post action="FormNewEmployee.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">New Employee Request Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
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

<!-- EMPLOYEE DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Employee Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Last Name: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=LastName></TD>
						</TR>
						<TR>
							<TD><B>Middle Names:</B></TD>
							<TD><INPUT name=MiddleNames></TD>
						</TR>
						<TR>
							<TD><B>First Name: <SPAN class="Warning">*</SPAN></B></TD>
							<TD><INPUT name=FirstName></TD>
						</TR>
						<TR>
							<TD><B>Date of Birth:</B></TD>
							<TD><INPUT name=DOB onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>

<!-- EMPLOYMENT DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Employment Details</TD></TR>
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
							<TD><B>Direct Supervisor:</B></TD>
							<TD><INPUT name=DirectSupervisor></TD>
						</TR>
						<TR>
							<TD><B>Cost Centre:</B></TD>
							<TD><INPUT name=CostCentre></TD>
						</TR>
						<TR>
							<TD><B>Position Title:</B></TD>
							<TD><INPUT name=PositionTitle></TD>
						</TR>
						<TR>
							<TD><B>Contract Start Date:</B></TD>
							<TD><INPUT name=StartDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR>
							<TD><B>Contract End Date:</B></TD>
							<TD><INPUT name=EndDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR>
							<TD><B>User has prior Network Access?</B></TD>
							<TD>
								<SELECT name=PreviousNetwork onchange="ShowHide(this, 'PREVNETWORK')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR id=PREVNETWORK style="DISPLAY:None">
							<TD id=PREVNETWORK><B>Previous User ID:</B></TD>
							<TD><INPUT name=PrevUserID></TD>
						</TR>
						<TR>
							<TD><B>Name of Contractor Company:</B></TD>
							<TD>
								<INPUT name=ContractorCompany>
							</TD>
						</TR>

<!-- POSITION DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Position Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Existing Position?</B></TD>
							<TD>
								<SELECT name=ExistingPosition onchange="ShowHide(this, 'EXISTINGPOS')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR id=EXISTINGPOS style="DISPLAY:None">
							<TD id=EXISTINGPOS><B>Who are they replacing?</B></TD>
							<TD><INPUT name=WhoReplacing title="Specify another User Name"></TD>
						</TR>
						<TR>
							<TD><B>Position Number (if known):</B></TD>
							<TD><INPUT name=PositionNumber></TD>
						</TR>
						<TR>
							<TD><B>Same roster as:</B></TD>
							<TD><INPUT name=SameRosterAs title="Specify another User Name"></TD>
						</TR>
						
<!-- NETWORK ACCESS DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Network Access Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Is Network Access required? <SPAN class="Warning">*</SPAN></B></TD>
							<TD>
								<SELECT name=NetworkAccess onchange="ShowHide(this, 'NETWORK')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR ID=NETWORK style="DISPLAY:None">
							<TD><B>Internet Access Required?</B></TD>
							<TD>
								<SELECT name=Internet>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR ID=NETWORK style="DISPLAY:None">
							<TD><B>Dial In Access Required?</B></TD>
							<TD>
								<SELECT name=DialIn>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR ID=NETWORK style="DISPLAY:None">
							<TD><B>Provide same network access as:</B></TD>
							<TD><INPUT name=SameNetworkAccess title="Specify another User Name"></TD>
						</TR>
						<TR ID=NETWORK style="DISPLAY:None">
							<TD><B>Specify Folders Required:</B></TD>
							<TD><INPUT name=FoldersRequired title="eg N:\(projects) & G:\(IT)"></TD>
						</TR>
						<TR ID=NETWORK style="DISPLAY:None">
							<TD><B>Provide same email groups as:</B></TD>
							<TD><INPUT name=EmailGroups title="Specify another User Name"></TD>
						</TR>
						<TR ID=NETWORK style="DISPLAY:None">
							<TD><B>Access To Other Mailbox:</B></TD>
							<TD><INPUT name=OtherMailbox title="Specify Mailbox Name"></TD>
						</TR>
						<TR ID=NETWORK style="DISPLAY:None">
							<TD><B>How will the user connect?</B></TD>
							<TD>
								<SELECT name=Connect>
									<OPTION value="">-- Select a Connection Type --</OPTION>
									<OPTION value="Oxiana Supplied PC/Notebook">Oxiana Supplied PC/Notebook</OPTION>
									<OPTION value="Their own PC/Notebook">Their own PC/Notebook</OPTION>
									<OPTION value="Both">Both</OPTION>
								</SELECT>
							</TD>
						</TR>

<!-- ELLIPSE ACCESS DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Ellipse Access Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Is Ellipse Access Required? <SPAN class="Warning">*</SPAN></B></TD>
							<TD>
								<SELECT name=Ellipse onchange="ShowHide(this, 'ELLIPSE')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR ID=ELLIPSE style="DISPLAY:None">
							<TD><B>Ellipse Default Printer:</B></TD>
							<TD><INPUT name=EllipsePrinter></TD>
						</TR>
						<TR ID=ELLIPSE style="DISPLAY:None"><TD>&nbsp;</TD></TR>
						<TR ID=ELLIPSE style="DISPLAY:None">
							<TD valign=top><B>District Access Required:</B></TD>
							<TD>
								<INPUT type=checkbox name=6000 value="6000 - Melbourne"><SPAN style="WIDTH:125px">6000 - Melbourne</SPAN>
								<INPUT type=checkbox name=6030 value="6030 - Golden Grove"><SPAN style="WIDTH:125px">6030 - Golden Grove</SPAN><BR>
								<INPUT type=checkbox name=6040 value="6040 - Jundee"><SPAN style="WIDTH:125px">6040 - Jundee</SPAN>
								<INPUT type=checkbox name=6090 value="6090 - Tanami"><SPAN style="WIDTH:125px">6090 - Tanami</SPAN><BR>
								<INPUT type=checkbox name=6530 value="6530 - Adelaide & Perth"><SPAN style="WIDTH:125px">6530 - Adelaide & Perth</SPAN>
								<INPUT type=checkbox name=6880 value="6880 - Waihi"><SPAN style="WIDTH:125px">6880 - Waihi</SPAN><BR>
								<INPUT type=checkbox name=6950 value="6950 - Pajingo"><SPAN style="WIDTH:125px">6950 - Pajingo</SPAN>
								<INPUT type=checkbox name=ALL value="ALL - All Districts"><SPAN style="WIDTH:125px">ALL - All Districts</SPAN>
							</TD>
						</TR>
						<TR ID=ELLIPSE style="DISPLAY:None"><TD>&nbsp;</TD></TR>
						<TR ID=ELLIPSE style="DISPLAY:None">
							<TD valign=top><B>Same Access As:</B></TD>
							<TD>
								<INPUT name=EllipseSameAccessAs>
							</TD>
						</TR>

<!-- AHFM ACCESS DETAILS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Australian Hyperion Financial Management (AHFM)</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Is AHFM Access Required? <SPAN class="Warning">*</SPAN></B></TD>
							<TD>
								<SELECT name=AHFM onchange="ShowHide(this, 'AHFM')">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR ID=AHFM style="DISPLAY:None">
							<TD><B>Provide same AHFM access as:</B></TD>
							<TD><INPUT name=AHFMAccess title="Specify another User Name"></TD>
						</TR>
						<TR ID=AHFM style="DISPLAY:None">
							<TD><B>Hyperion Reporting Required?</B></TD>
							<TD>
								<SELECT name=HyperionReporting>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR ID=AHFM style="DISPLAY:None"><TD>&nbsp;</TD></TR>
						<TR ID=AHFM style="DISPLAY:None">
							<TD><B>Group access required for AHFM:</B></TD>
							<TD>
								<INPUT type=checkbox name=AHFMAccess1 value="Admin - Application"><SPAN style="WIDTH:125px">Admin - Application</SPAN>
								<INPUT type=checkbox name=AHFMAccess2 value="Admin - Process Mgmt"><SPAN style="WIDTH:125px">Admin - Process Mgmt</SPAN><BR>
								<INPUT type=checkbox name=AHFMAccess3 value="Admin - Security"><SPAN style="WIDTH:125px">Admin - Security</SPAN>
								<INPUT type=checkbox name=AHFMAccess4 value="Golden Grove"><SPAN style="WIDTH:125px">Golden Grove</SPAN><BR>
								<INPUT type=checkbox name=AHFMAccess5 value="Jundee"><SPAN style="WIDTH:125px">Jundee</SPAN>
								<INPUT type=checkbox name=AHFMAccess6 value="KCGM"><SPAN style="WIDTH:125px">KCGM</SPAN><BR>
								<INPUT type=checkbox name=AHFMAccess7 value="Tanami"><SPAN style="WIDTH:125px">Tanami</SPAN>
								<INPUT type=checkbox name=AHFMAccess8 value="Pajingo"><SPAN style="WIDTH:125px">Pajingo</SPAN><BR>
								<INPUT type=checkbox name=AHFMAccess9 value="Regional"><SPAN style="WIDTH:125px">Regional</SPAN>
								<INPUT type=checkbox name=AHFMAccess10 value="Supervisor"><SPAN style="WIDTH:125px">Supervisor</SPAN><BR>
								<INPUT type=checkbox name=AHFMAccess11 value="Manager"><SPAN style="WIDTH:125px">Manager</SPAN>
								<INPUT type=checkbox name=AHFMAccess12 value="View Only"><SPAN style="WIDTH:125px">View Only</SPAN>
							</TD>
						</TR>

<!-- OTHER APPLICATIONS -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Other Applications</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=Warning>Please note: Additional approval forms may be required for each application.</TD></TR>
						<TR><TD>&nbsp;</TD><TD><B>Profile/Type of Access</B></TD></TR>
						<TR>
							<TD><B>ANZ Online:</B></TD>
							<TD><INPUT name=ANZOnline></TD>
						</TR>
						<TR>
							<TD><B>Cognos ePlanning:</B></TD>
							<TD><INPUT name=CognosePlanning></TD>
						</TR>
						<TR>
							<TD><B>DirectLink (NBNZ Online - Adelaide):</B></TD>
							<TD><INPUT name=DirectLinkAdelaide></TD>
						</TR>
						<TR>
							<TD><B>DirectLink (NBNZ Online - Waihi):</B></TD>
							<TD><INPUT name=DirectLinkWaihi></TD>
						</TR>
						<TR>
							<TD><B>Hyperion Enterprise (HE):</B></TD>
							<TD><INPUT name=HE></TD>
						</TR>
						<TR>
							<TD><B>Hyperion Planning (HP):</B></TD>
							<TD><INPUT name=HP></TD>
						</TR>
						<TR>
							<TD><B>SAP Read Only Access:</B></TD>
							<TD><INPUT name=SAP></TD>
						</TR>

<!-- ADDITIONAL INFORMATION -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Additional Information</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Specify any other details which could not be recorded above:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=OtherDetails style="WIDTH:450px;HEIGHT:75px"></TEXTAREA></TD></TR>

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