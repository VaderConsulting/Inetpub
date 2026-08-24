<%@Language=VBScript%>
<!-- #Include file="Includes/Menu.asp" -->
<!-- #Include file="Includes/FormatDate.asp" -->
<%
	Response.Expires = 0
	
	'check if the user has clicked the submit button
	If Request.Form("PostBack") = "true" Then
		
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
		lstrHTML = lstrHTML & ".Box" & vbcrlf
		lstrHTML = lstrHTML & "{" & vbcrlf
		lstrHTML = lstrHTML & " BORDER: 1px solid gray;" & vbcrlf
		lstrHTML = lstrHTML & " HEIGHT: 30px;" & vbcrlf
		lstrHTML = lstrHTML & "}" & vbcrlf
		lstrHTML = lstrHTML & "</STYLE>" & vbcrlf
		lstrHTML = lstrHTML & "</HEAD>" & vbcrlf
		lstrHTML = lstrHTML & "<BODY>" & vbcrlf
		lstrHTML = lstrHTML & "	<TABLE cellpadding=0 cellspacing=0 width=570>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2>Project Initiation Request Form</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		
'PROJECT INFORMATION
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Project Information</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Project Number:</B></TD><TD>" & Request.Form("ProjectNumber") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Title:</B></TD><TD>" & Request.Form("Title") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Initiated By:</B></TD><TD>" & Request.Form("InitiatedBy") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Approved By:</B></TD><TD>" & Request.Form("ApprovedBy") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Project Owner:</B></TD><TD>" & Request.Form("ProjectOwner") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Scheduled Commencement Date:</B></TD><TD>" & Request.Form("CommencementDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Scheduled Completion Date:</B></TD><TD>" & Request.Form("CompletionDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"

'PROJECT DEFINITION
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Project Information</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Objectives / Scope:</B></TD><TD>" & Replace(Request.Form("Scope"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Deliverables:</B></TD><TD>" & Replace(Request.Form("Deliverables"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Exclusions / Constraints:</B></TD><TD>" & Replace(Request.Form("Exclusions"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Assumptions:</B></TD><TD>" & Replace(Request.Form("Assumptions"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Hardware / Software Required:</B></TD><TD>" & Replace(Request.Form("Hardware"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Project Prerequisites:</B></TD><TD>" & Replace(Request.Form("Prerequisites"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "		<TR><TD valign=top><B>Costs:</B></TD><TD>" & Replace(Request.Form("Cost"), vbcrlf, "<BR>") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>"
		
'AUTHORISATION OF WORK
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Authorisation of Work</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD colspan=2><TABLE cellpadding=0 cellspacing=0 width=570>" & vbcrlf
		lstrHTML = lstrHTML & "			<TR><TH width=195></TH><TH width=100>Name</TH><TH width=100>Title</TH><TH width=100>Signature</TH><TH width=75>Date</TH></TR>" & vbcrlf
		lstrHTML = lstrHTML & "			<TR><TD><B>Oxiana Representative</B></TD><TD class=Box>" & Request.Form("OXName") & "&nbsp;</TD><TD class=Box>" & Request.Form("OXTitle") & "&nbsp;</TD><TD class=Box>" & Request.Form("OXSignature") & "&nbsp;</TD><TD class=Box>" & Request.Form("OXDate") & "&nbsp;</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "			<TR><TD><B>Empired Representative</B></TD><TD class=box>" & Request.Form("EMPName") & "&nbsp;</TD><TD class=Box>" & Request.Form("EMPTitle") & "&nbsp;</TD><TD class=Box>" & Request.Form("EMPSignature") & "&nbsp;</TD><TD class=Box>" & Request.Form("EMPDate") & "&nbsp;</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		</TABLE></TD></TR>"
										
		lstrHTML = lstrHTML & "			</TD></TR>"
		lstrHTML = lstrHTML & "		</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'Response.Write lstrHTML
		'Response.End
		
		'now write the file
		Dim OutputPath
		OutputPath = Server.MapPath(Application("Documents") & "Forms/ProjectInitiation")
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
		lstrTo = Application("EmailProjectInitiation")
		lstrSubject = EmailSubjectDev & "Project Initiation Request"
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
	<TITLE>IT Intranet: Project Initiation Request Form</TITLE>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Project Initiation Request</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>Project Initiation Request Submitted</P>
				<P>Your request has been emailed to the IT Helpdesk management for review.</P>
				<DIV class=Copyright><BR>This page last updated: 04 Oct 2006 (Empired)</DIV>
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
	<TITLE>IT Intranet: Project Initiation Request</TITLE>
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
				<FORM method=post action="FormProjectInitiation.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Project Initiation Request Form</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
							</TD>
						</TR>
						
<!-- PROJECT INFORMATION -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD class=SubHeading>Project Information</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Project Number:</B></TD>
							<TD><INPUT name=ProjectNumber><INPUT type=hidden name=PostBack value="true"></TD>
						</TR>
						<TR>
							<TD><B>Title:</B></TD>
							<TD><INPUT name=Title></TD>
						</TR>
						<TR>
							<TD><B>Initiated By:</B></TD>
							<TD><INPUT name=InitiatedBy></TD>
						</TR>
						<TR>
							<TD><B>Approved By:</B></TD>
							<TD><INPUT name=ApprovedBy></TD>
						</TR>
						<TR>
							<TD><B>Project Owner:</B></TD>
							<TD><INPUT name=ProjectOwner></TD>
						</TR>
						<TR>
							<TD><B>Scheduled Commencement Date:</B></TD>
							<TD><INPUT name=CommencementDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR>
							<TD><B>Scheduled Completion Date:</B></TD>
							<TD><INPUT name=CompletionDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						
<!-- PROJECT DEFINITION -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD class=SubHeading>Project Definition</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Objectives / Scope:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=Scope style="WIDTH:570px;HEIGHT:50px"></TEXTAREA></TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Deliverables:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=Deliverables style="WIDTH:570px;HEIGHT:50px"></TEXTAREA></TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Exclusions / Constraints:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=Exclusions style="WIDTH:570px;HEIGHT:50px"></TEXTAREA></TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Assumptions:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=Assumptions style="WIDTH:570px;HEIGHT:50px"></TEXTAREA></TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Hardware / Software Required:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=Hardware style="WIDTH:570px;HEIGHT:50px"></TEXTAREA></TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Project Prerequisites:</B></TD></TR>
						<TR><TD colspan=2><TEXTAREA name=Prerequisites style="WIDTH:570px;HEIGHT:50px"></TEXTAREA></TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Costs:</B></TD>
							<TD><INPUT type=radio name=Cost value="Fixed Price"> Fixed Price&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=radio name=Cost value="Time and Materials"> Time and Materials</TD>
						</TR>
						<TR><TD colspan=2 class=Warning>NOTE: Fixed Price costs will not be broken down item by item; Time and materials total fixed price is based on estimated hours only.</TD></TR>
						
<!-- AUTHORISATION OF WORK -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD class=SubHeading>Authorisation of Work</TD></TR>
						<TR>
							<TD colspan=2>
								<TABLE cellpadding=0 cellspacing=0 width=570>
									<TR>
										<TD width=195></TD>
										<TD width=100>Name</TD>
										<TD width=100>Title</TD>
										<TD width=100>Signature</TD>
										<TD width=75>Date</TD>
								   </TR>
								   <TR>
										<TD><B>Oxiana Representative</B></TD>
										<TD><INPUT name=OXName style="WIDTH:100px" value="<%= lstrUserFullName%>"></TD>
										<TD><INPUT name=OXTitle style="WIDTH:100px"></TD>
										<TD><INPUT name=OXSignature style="WIDTH:100px"></TD>
										<TD><INPUT name=OXDate onblur="CheckDate(this, true)" style="WIDTH:75px" value="<%= FormatDate(Now, C_FormatDate_ShortDate)%>"></TD>
								   </TR>
								   <TR>
										<TD><B>Empired Representative</B></TD>
										<TD><INPUT name=EMPName style="WIDTH:100px"></TD>
										<TD><INPUT name=EMPTitle style="WIDTH:100px"></TD>
										<TD><INPUT name=EMPSignature style="WIDTH:100px"></TD>
										<TD><INPUT name=EMPDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
								   </TR>
								</TABLE>
							</TD>
						</TR>

<!-- SUBMIT -->
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 align=center><INPUT type=button value=Submit onclick="form.submit()"></TD></TR>
					</TABLE>
				</FORM>
				<DIV class=Copyright><BR>This page last updated: 04 Oct 2006 (Empired)</DIV>
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