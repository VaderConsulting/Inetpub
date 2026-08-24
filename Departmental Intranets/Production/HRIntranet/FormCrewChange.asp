<%@  language="VBScript" %>
<!-- #Include file="Includes/Menu.asp" -->
<!-- #Include file="Includes/FormatDate.asp" -->
<%
	response.expires = 0

'this subr/outine will save the passed in string to the specified file
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
	If Len(Request.Form("Surname")) > 0 Then
		
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
		lstrSubject = EmailSubjectDev & "Position Title, Reporting Line or Crew Change Request"
		lstrServer =  Application("EmailSMTPServer")
		
		lstrCC = Session("EmailAddr")
		if Len(Trim(Request.form("EmailCC"))) > 0 Then
			lstrCC = Trim(Request.Form("EmailCC"))
		End If

		lstrHTML = lstrHTML & "<HTML>" & vbcrlf
		lstrHTML = lstrHTML & "<head>" & vbcrlf
		lstrHTML = lstrHTML & "<STYLE>" & vbcrlf
		lstrHTML = lstrHTML & "BODY, table" & vbcrlf
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
		lstrHTML = lstrHTML & "</head>" & vbcrlf
		lstrHTML = lstrHTML & "<body>" & vbcrlf
		lstrHTML = lstrHTML & "	<table cellpadding=0 cellspacing=0 width=570>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td class=MainHeading colspan=2>Crew Change Request</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td width=200>&nbsp;</td><td width=370></td></tr>" & vbcrlf
		
		'now build the HTML to send to the help-desk
		
		'EMPLOYEE DETAILS
		lstrHTML = lstrHTML & "		<tr><td colspan=2 class=SubHeading>Employee Details</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Surname:</b></td><td>" & Request.Form("Surname") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Given Names:</b></td><td>" & Request.Form("GivenNames") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Site:</b></td><td>" & Request.Form("Site") & "</td></tr>" & vbcrlf
		
		lstrHTML = lstrHTML & "		<tr><td><b>Position Number:</b></td><td>" & Request.Form("PositionNo") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Current Supervisor's Title:</b></td><td>" & Request.Form("SuperTitle") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Current Supervisor's Name:</b></td><td>" & Request.Form("SuperName") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Current Position Title:</b></td><td>" & Request.Form("PosTitle") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Revised Position Title:</b></td><td>" & Request.Form("NewPosTitle") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Job:</b></td><td>" & Request.Form("Job") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>New Supervisor's Title:</b></td><td>" & Request.Form("NewSuperTitle") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>New Supervisor's Name:</b></td><td>" & Request.Form("NewSuperName") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Effective Date of Changes:</b></td><td>" & Request.Form("EffectiveDate") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>New Cost Centre:</b></td><td>" & Request.Form("NewCostCentre") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td colspan=2>NOTE: If the employee is changing crews please complete the roster information box below:</td></tr>" & vbCrLf
							
		lstrHTML = lstrHTML & "		<tr><td><b>The employee will work the same roster as:</b></td><td>" & Request.Form("SameRosterAs") & "</td></tr>" & vbcrlf
		lstrHTML = lstrHTML & "		<tr><td><b>Comments:</b></td><td>" & Request.Form("Comments") & "</td></tr>" & vbcrlf
		
		lstrHTML = lstrHTML & "	</table>" & vbcrlf
		lstrHTML = lstrHTML & "	</body>" & vbcrlf
		lstrHTML = lstrHTML & "</html>"
		
		'save the file
		SaveToFile Server.MapPath(Application("Documents") & "Forms\CrewChange") & "\" & FormatDate(Now, C_FormatDate_Identifier) & ".html", lstrHTML
		
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
<html>
<head>
    <title>HR Intranet: Position Title, Reporting Line or Crew Change</title>
    <link rel="stylesheet" href="Styles/Style.css" type="text/css" />
</head>
<body>
    <table width="770" cellpadding="0" cellspacing="0" align="center" width="770">
        <tr id="HidePrint" style="height:79;">
            <td colspan="3" background="Images/Logo.gif" align="right" class="Copyright">
                <a href="SiteMap.asp">Site Map</a><img align="absmiddle" width="19" height="12" src="Images/Privacydivider.gif"><a href="ContactUs.asp">Contact Us</a>
            </td>
        </tr>
        <tr id="HidePrint">
            <td colspan="3">
                <img height="11" width="770" src="Images/DottedLine.gif" />
            </td>
        </tr>
        <tr id="HidePrint">
            <td width="190">
                <div class="HomeLink">
                    <a href="index.asp" width="100%">HOME</a></div>
            </td>
            <td width="10">
                &nbsp;
            </td>
            <td width="570">
                &nbsp;
            </td>
        </tr>
        <tr id="HidePrint">
            <td colspan="3">
                <img height="11" width="770" src="Images/DottedLine.gif">
            </td>
        </tr>
        <tr>
            <td id="HidePrint" width="190" valign="top">
                <table cellpadding="0" cellspacing="0" width="100%">
                    <%ShowMenu%>
                </table>
            </td>
            <td width="10">
            </td>
            <td width="570" valign="top">
                <span class="MainHeading" style="display: inline; width: 460px; height: 21px">Position Title, Reporting Line or Crew Change Request</span><span class="Copyright" style="padding-top: 6px; padding-right: 25px; display: inline; width: 110px; height: 21px; cursor: hand; background-image: url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</span>
                <p class="SubHeading">
                    Position Title, Reporting Line or Crew Change Request Submitted</p>
                <p>
                    Your request has been emailed to the IT Helpdesk. You will be contacted by one of the Helpdesk staff as soon as possible.</p>
                <div class="Copyright">
                    <br />
                    This page last updated: 06 Apr 2008 (Empired)</div>
            </td>
        </tr>
        <tr id="HidePrint">
            <td colspan="3">
                <img height="11" width="770" src="Images/DottedLine.gif">
            </td>
        </tr>
        <tr id="HidePrint">
            <td colspan="3" class="Copyright">
                <a href="IntranetDisclaimer.asp">Disclaimer</a><img align="absmiddle" width="19" height="12" src="Images/Privacydivider.gif"><a href="IntranetPrivacyStatement.asp">Privacy</a><img align="absmiddle" width="19" height="12" src="Images/Privacydivider.gif">Copyright Oxiana Limited 2006
            </td>
        </tr>
    </table>
</body>
</html>
<%		
	Else

		lstrUserFullName = Session("SAMUserName")
		lstrEmail = Session("EmailAddr")
%>
<html>
<head>
    <title>HR Intranet: Position Title, Reporting Line or Crew Change Request</title>
    <link rel="stylesheet" href="Styles/Style.css" type="text/css" />

    <script type="text/javascript" language="javascript">
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
    </script>

    <script type="text/javascript" language="javascript">
	<!--
	function ValidateAndSubmit()
	{
		var TheForm = document.forms.Form1;
		
		if (CheckField(TheForm.Surname, '', 'Please specify the users Surname') == false) return false;
		if (CheckField(TheForm.GivenNames, '', 'Please specify the users Given Names') == false) return false;
		if (CheckField(TheForm.Site, '', 'Please specify the users Site') == false) return false;
		if (CheckField(TheForm.SuperTitle, '', 'Please specify the Supervisors title') == false) return false;
		if (CheckField(TheForm.SuperName, '', 'Please specify the Supervisors name') == false) return false;
		if (CheckField(TheForm.EffectiveDate, '', 'Please specify the Effective Date of changes') == false) return false;
		
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
    </script>

    <script type="text/javascript" language="javascript" src="Scripts/HTMLDateFormat.js"></script>

</head>
<body onbeforeprint="HideNonPrintables()" onafterprint="ShowNonPrintables()">
    <table width="770" cellpadding="0" cellspacing="0" align="center" width="770">
        <tr id="HidePrint" height="79">
            <td colspan="3" background="Images/Logo.gif" align="right" class="Copyright">
                <a href="SiteMap.asp">Site Map</a><img align="absmiddle" width="19" height="12" src="Images/Privacydivider.gif"><a href="ContactUs.asp">Contact Us</a>
            </td>
        </tr>
        <tr id="HidePrint">
            <td colspan="3">
                <img height="11" width="770" src="Images/DottedLine.gif">
            </td>
        </tr>
        <tr id="HidePrint">
            <td width="190">
                <div class="HomeLink">
                    <a href="index.asp" width="100%">HOME</a></div>
            </td>
            <td width="10">
                &nbsp;
            </td>
            <td width="570">
                &nbsp;
            </td>
        </tr>
        <tr id="HidePrint">
            <td colspan="3">
                <img height="11" width="770" src="Images/DottedLine.gif">
            </td>
        </tr>
        <tr>
            <td id="HidePrint" width="190" valign="top">
                <table cellpadding="0" cellspacing="0" width="100%">
                    <%ShowMenu%>
                </table>
            </td>
            <td width="10">
            </td>
            <td width="570" valign="top">
                <form id="Form1" method="post" action="FormCrewChange.asp">
                <table cellpadding="0" cellspacing="0" width="570" id="form">
                    <tr>
                        <td colspan="2">
                            <span class="MainHeading" style="display: inline; width: 460px; height: 21px">Position Title, Reporting Line or Crew Change Request</span><span class="Copyright" style="padding-top: 6px; padding-right: 25px; display: inline; width: 110px; height: 21px; cursor: hand; background-image: url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</span>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            Is this form being completed to change a position title?&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="PosChange" value="Yes">
                            Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="PosChange" value="No">
                            No
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            Is this form being completed to change a reporting line and/or crew change?<input type="radio" name="ReportingChange" value="Yes">
                            Yes&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="radio" name="ReportingChange" value="No">
                            No
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Surname: <span class="Warning">*</span></b>
                        </td>
                        <td>
                            <input name="Surname">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Given Names: <span class="Warning">*</span></b>
                        </td>
                        <td>
                            <input name="GivenNames">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Site/Location: <span class="Warning">*</span></b>
                        </td>
                        <td>
                            <select name="Site">
                                <option value="">-- Select a Site --</option>
                                <option value="Golden Grove">Golden Grove</option>
                                <option value="Melbourne">Melbourne</option>
                                <option value="Adelaide">Adelaide</option>
                                <option value="SPDG">SPDG</option>
                                <option value="Other">Other...</option>
                            </select>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Position Number:</b>
                        </td>
                        <td>
                            <input name="PositionNo">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Current Supervisor's Title: <span class="Warning">*</span></b>
                        </td>
                        <td>
                            <input name="SuperTitle">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Current Supervisor's Name: <span class="Warning">*</span></b>
                        </td>
                        <td>
                            <input name="SuperName">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Current Position Title:</b>
                        </td>
                        <td>
                            <input name="PosTitle">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Revised Position Title:</b>
                        </td>
                        <td>
                            <input name="NewPosTitle">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Job:</b>
                        </td>
                        <td>
                            <input name="Job">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>New Supervisor's Title:</b>
                        </td>
                        <td>
                            <input name="NewSuperTitle">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>New Supervisor's Name:</b>
                        </td>
                        <td>
                            <input name="NewSuperName" />
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Effective Date of Changes: <span class="Warning">*</span></b>
                        </td>
                        <td>
                            <input name="EffectiveDate" onblur="CheckDate(this, true)" style="width: 75px">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>New Cost Centre:</b>
                        </td>
                        <td>
                            <input name="NewCostCentre">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <b>NOTE: If the employee is changing crews please complete the roster information box below:</b>
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>The employee will work the same roster as:</b>
                        </td>
                        <td>
                            <input name="SameRosterAs">
                        </td>
                    </tr>
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <b>Comments:</b>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <textarea name="Comments" style="width: 450px; height: 75px"></textarea>
                        </td>
                    </tr>
                    <!-- EMAILCC -->
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td>
                            <b>Email CC:</b>
                        </td>
                        <td>
                            <input name="EmailCC" readonly="readonly" style="background-color: #E8E8E8; width: 200px"><input type="button" value="..." onclick="PromptForEmail()">
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            (If you wish to send this form to an additional email address, enter the email address here. Please ensure you enter the correct address.)
                        </td>
                    </tr>
                    <!-- SUBMIT -->
                    <tr>
                        <td>
                            &nbsp;
                        </td>
                    </tr>
                    <tr>
                        <td align="center" colspan="2">
                            <input type="button" value="Submit" onclick="ValidateAndSubmit()">
                        </td>
                    </tr>
                </table>
                </form>
                <div class="Copyright">
                    <br />
                    This page last updated: 6 Apr 2008 (Empired)</div>
            </td>
        </tr>
        <tr id="HidePrint">
            <td colspan="3">
                <img height="11" width="770" src="Images/DottedLine.gif"/>
            </td>
        </tr>
        <tr id="HidePrint">
            <td colspan="3" class="Copyright">
                <a href="IntranetDisclaimer.asp">Disclaimer</a><img align="absmiddle" width="19" height="12" src="Images/Privacydivider.gif"/><a href="IntranetPrivacyStatement.asp">Privacy</a><img align="absmiddle" width="19" height="12" src="Images/Privacydivider.gif">Copyright Oxiana Limited 2006
            </td>
        </tr>
    </table>
</body>
</html>
<%
	End If
%>