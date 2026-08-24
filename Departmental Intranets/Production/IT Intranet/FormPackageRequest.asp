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
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2>Application Request Form</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		'now build the HTML to send to the help-desk
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Requestor Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Submit Date:</B></TD><TD>" & Request.Form("SubmitDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Required By Date:</B></TD><TD>" & Request.Form("RequiredDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>User Account:</B></TD><TD>" & Request.Form("UserName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Requestor:</B></TD><TD>" & Request.Form("Requestor") & "</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Application Details</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Application Name:</B></TD><TD>" & request.form("AppName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Target Users:</B></TD><TD>" & request.form("TargetUsers") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Already Purchased?</B></TD><TD>" & request.form("AlreadyPurchased") & "</TD></TR>" & vbcrlf
		If Request.form("AlreadyPurchased") = "No" Then
			lstrHTML = lstrHTML & "		<TR><TD><B>Cost Per License:</B></TD><TD>" & request.form("CostPerLicense") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "		<TR><TD><B>Number of Licenses Required:</B></TD><TD>" & request.form("NumLicensesRequired") & "</TD></TR>" & vbcrlf
		End If
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Justification / Business Case:</B></TD><TD>" & Replace(request.form("Justification"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Packaging Requirements</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Estimated Number of Installs:</B></TD><TD>" & request.form("NumInstalls") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Requires Packaging?</B></TD><TD>" & request.form("RequiresPackaging") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Prerequisite Software Required:</B></TD><TD>" & request.form("Prerequisites") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Location of Install Files:</B></TD><TD>" & request.form("SourceLocation") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Install / Config Instructions:</B></TD><TD>" & Replace(request.form("InstallInstructions"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		
		lstrHTML = lstrHTML & "		</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'now write the file
		Dim OutputPath
		OutputPath = Server.MapPath(Application("Documents") & "Forms/ApplicationPackageRequest")
		Dim FileName
		FileName = "\" & FormatDate(Now, C_FormatDate_Identifier) & " " & Request.Form("AppName") & ".html"
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
		lstrTo = Application("EmailApplicationPackageRequest")
		lstrSubject = EmailSubjectDev & "Application Package Request"
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
	<TITLE>IT Intranet: Application Request Form</TITLE>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Application Request Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>Application Request Submitted</P>
				<P>Your request has been emailed to the IT Helpdesk management for review.</P>
				<DIV class=Copyright><BR>This page last updated: 20 Sep 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>		
<%		
	Else
		
		lstrUserFullName = Session("UsersName")
		lstrEmail = Session("EmailAddr")
%>
<HTML>
<HEAD>
	<TITLE>IT Intranet: Application Request Form</TITLE>
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
	function ComboChange(TheCombo)
	{
		if (TheCombo.value == 'No')
		{
			CPL.style.display = 'Inline';
			NLR.style.display = 'Inline';
			ESTIMATED.style.display = 'None';
		}
		else
		{
			CPL.style.display = 'None';
			NLR.style.display = 'None';
			ESTIMATED.style.display = 'Inline';
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
				<FORM method=post action="FormPackageRequest.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Application Request Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
							</TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Requester Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD width=200><B>Submit Date:</B></TD>
							<TD width=370><INPUT name=SubmitDate value="<%= FormatDate(Now, C_FormatDate_ShortDate)%>" onblur="CheckDate(this, true)" style="WIDTH:75px" readonly></TD>
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
						<TR><TD colspan=2 class=SubHeading>Application Details</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Application Name:</B></TD>
							<TD><INPUT name=AppName style="WIDTH:250px"> <INPUT type=button value="View Software" onclick="window.open('<%= Application("Documents")%>XML/Software.xml', 'OXSoftware', 'SCROLLBARS=yes,HEIGHT=400,WIDTH=570')"></TD>
						</TR>
						<TR>
							<TD><B>Target Users:</B></TD>
							<TD><INPUT name=TargetUsers style="WIDTH:370px"></TD>
						</TR>
						<TR>
							<TD><B>Already Purchased?</B></TD>
							<TD>
								<SELECT name=AlreadyPurchased onchange="ComboChange(this)">
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR id=ESTIMATED style="DISPLAY:None">
							<TD>&nbsp;</TD>
							<TD><SPAN style="WIDTH:200px"><B>Estimated Cost (if known):</B></SPAN><INPUT name=EstimatedCost style="WIDTH:170px"></TD>
						</TR>
						<TR id=CPL style="DISPLAY:None">
							<TD>&nbsp;</TD>
							<TD><SPAN style="WIDTH:200px"><B>Cost Per License:</B></SPAN><INPUT name=CostPerLicense style="WIDTH:170px"></TD>
						</TR>
						<TR id=NLR style="DISPLAY:None">
							<TD>&nbsp;</TD>
							<TD><SPAN style="WIDTH:200px"><B>Number of Licenses Required:</B></SPAN><INPUT name=NumLicensesRequired style="WIDTH:170px"></TD>
						</TR>
						
						<TR><TD colspan=2><B>Justification / Business Case:</B></TD></TR>
						<TR>
							<TD colspan=2><TEXTAREA name=Justification style="WIDTH:570px;HEIGHT:75px"></TEXTAREA></TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Packaging Requirements</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Estimated Number of Installs:</B></TD>
							<TD><INPUT name=NumInstalls style="WIDTH:70px"></TD>
						</TR>
						<TR>
							<TD><B>Requires Packaging?</B></TD>
							<TD>
								<SELECT name=RequiresPackaging>
									<OPTION value="">-- Select --</OPTION>
									<OPTION value="Yes">Yes</OPTION>
									<OPTION value="No">No</OPTION>
								</SELECT>
							</TD>
						</TR>
						<TR>
							<TD><B>Prerequisite Software Required?</B></TD>
							<TD><INPUT name=Prerequisites style="WIDTH:370px"></TD>
						</TR>
						<TR>
							<TD><B>Location of Install Files:</B></TD>
							<TD><INPUT name=SourceLocation style="WIDTH:370px"></TD>
						</TR>
						<TR>
							<TD colspan=2><B>Install/Config instructions:</B></TD>
						</TR>
						<TR>
							<TD colspan=2><TEXTAREA name=InstallInstructions style="WIDTH:570px;HEIGHT:75px"></TEXTAREA></TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 align=center><INPUT type=button value=Submit onclick="form.submit()"></TD></TR>
					</TABLE>
				</FORM>
				<DIV class=Copyright><BR>This page last updated: 20 Sep 2006 (Empired)</DIV>
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