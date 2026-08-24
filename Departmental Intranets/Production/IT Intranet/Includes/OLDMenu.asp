<%
'define the variables
Dim PageID

CheckDevelopmentSite

private sub CheckDevelopmentSite()
	
	'check if this request has come from the dev site
	If Right(LCase(Request.Servervariables("Server_name")), 3) = "dev" Then
		Response.write "<div class=""MainHeading"" align=center>THIS IS A DEVELOPMENT SITE</div>"
		Response.write "<div align=center class=Warning>This site is subject to development which could cause erratic behaviour. Please use the production site for all official work requests: <SPAN class=""Copyright""><A href=""HTTP://" & Left(Request.ServerVariables("Server_Name"), Len(Request.ServerVariables("Server_Name")) - 3) & """>" & Left(Request.ServerVariables("Server_Name"), Len(Request.ServerVariables("Server_Name")) - 3) & "</A></SPAN></div>"
	Elseif Instr(1, LCase(Request.ServerVariables("Server_Name")), "dev.") > 0 Then
		Response.write "<div class=""MainHeading"" align=center>THIS IS A DEVELOPMENT SITE</div>"
		Response.write "<div align=center class=Warning>This site is subject to development which could cause erratic behaviour. Please use the production site for all official work requests: <SPAN class=""Copyright""><A href=""HTTP://" & Replace(LCase(Request.ServerVariables("Server_Name")), "dev.", ".") & """>" & Replace(LCase(Request.ServerVariables("Server_Name")), "dev.", ".") & "</A></SPAN></div>"
	End If
End Sub

'this function will check if the current page is on a development server
'and will return an appropriate string to add to the subject line of an email
'
'OUTPUT:
'	String	- Will be blank if the site is not dev, otherwise "DEVELOPMENT SITE: " is returned
Function EmailSubjectDev()
	
	'check if this request has come from the dev site
	If (Right(LCase(Request.Servervariables("Server_name")), 3) = "dev") Or (Instr(1, LCase(Request.ServerVariables("Server_Name")), "dev.") > 0) Then
		EmailSubjectDev = "DEVELOPMENT SITE: "
	End If
End Function

Function Check(ThisPage)
	If PageID = ThisPage Then
		Check = "Selected"
	Else
		Check = ""
	End If
End Function

Sub ShowMenu()
	
	'define the local variables
	Dim llngLastPos
	Dim llngNextPos
	Dim TheURL
	
	'determine the URL of the page
	TheURL = UCase(Request.ServerVariables("URL"))
	
	llngLastPos = 1
	While llngLastPos > 0
		llngNextPos = llngLastPos
		llngLastPos = Instr(llngLastPos + 1, TheURL, "/")
	Wend
	
	TheURL = Right(TheURL, Len(TheURL) - llngNextPos)
	
	'determine the page ID of the current page
	Select Case TheURL
		Case "INDEX.ASP":					PageID = "000"
		
		Case "SERVICECENTRE.ASP":			PageID = "100"
		Case "HELPDESKSERVICES.ASP":		PageID = "101"
		Case "USERMANUALS.ASP":				PageID = "102"
		Case "TIPSANDTRICKS.ASP":			PageID = "103"
		Case "MEETHELPDESK.ASP":			PageID = "104"
		Case "FAQS.ASP":					PageID = "105"
		Case "HOTDESK.ASP":					PageID = "106"
		Case "SOFTWARELIST.ASP":					PageID = "107"

		Case "PROCUREMENT.ASP":				PageID = "200"
		Case "PROCUREMENTFAQS.ASP":			PageID = "201"
		Case "PROCUREMENTFORM.ASP":			PageID = "202"
		Case "HARDWARELIST.ASP":			PageID = "203"
		
		Case "TRAINING.ASP":				PageID = "300"
		Case "TRAINMANUALS.ASP":			PageID = "301"
		Case "TRAINLIVEMEETING.ASP":		PageID = "302"

		Case "POLICIESANDSTANDARDS.ASP":	PageID = "400"
		Case "POLICYACCEPTABLEUSAGE.ASP":	PageID = "401"
		Case "POLICYREMOTEACCESS.ASP":		PageID = "402"
		Case "POLICYSECURITY.ASP":			PageID = "403"
		Case "POLICYPROCUREMENT.ASP":		PageID = "404"
		Case "POLICYUSERACCOUNTS.ASP":		PageID = "405"

		Case "SELFSERVICE.ASP":				PageID = "500"
		Case "UPDATEDETAILS.ASP":			PageID = "501"
		Case "LOGSERVICECALL.ASP":			PageID = "502"
		
		Case "SITEMAP.ASP":					PageID = "600"
		Case "CONTACTUS.ASP":				PageID = "601"
		Case "INTRANETDISCLAIMER.ASP":		PageID = "602"
		Case "INTRANETPRIVACYSTATEMENT.ASP":	PageID = "603"
		
		Case "FORMS.ASP":					PageID = "800"
		Case "FORMITWORKREQUEST.ASP":		PageID = "801"
		Case "FORMPACKAGEREQUEST.ASP":		PageID = "802"
		Case "FORMPROJECTINITIATION.ASP":	PageID = "803"
		
		Case "PROJECTS.ASP":				PageID = "700"

		Case Else: Response.write TheURL
	End Select


%>
<TR class=MenuItem<%=Check("100")%>><TD><A href="ServiceCentre.asp">Help Desk</A></TD></TR>
<%If Left(PageID, 1) = "1" Then%>
	<TR class=SubMenuItem<%=Check("101")%>><TD><A href="HelpDeskServices.asp">Help Desk Services</A></TD></TR>
	<TR class=SubMenuItem<%=Check("102")%>><TD><A href="UserManuals.asp">User Manuals</A></TD></TR>
	<TR class=SubMenuItem<%=Check("103")%>><TD><A href="TipsAndTricks.asp">Tips And Tricks</A></TD></TR>
	<TR class=SubMenuItem<%=Check("104")%>><TD><A href="MeetHelpDesk.asp">Meet the Help Desk</A></TD></TR>
	<TR class=SubMenuItem<%=Check("105")%>><TD><A href="FAQs.asp">Frequently Asked Questions</A></TD></TR>
	<TR class=SubMenuItem<%=Check("106")%>><TD><A href="HotDesk.asp">Hot Desk</A></TD></TR>
	<TR class=SubMenuItem<%=Check("107")%>><TD><A href="SoftwareList.asp">Software Catalogue</A></TD></TR>
	
<%End If%>
<TR class=MenuItem<%=Check("200")%>><TD><A href="Procurement.asp">Procurement</A></TD></TR>
<%If Left(PageID, 1) = "2" Then%>
	<TR class=SubMenuItem<%=Check("201")%>><TD><A href="ProcurementFAQs.asp">Procurement FAQ's</A></TD></TR>
	<TR class=SubMenuItem<%=Check("202")%>><TD><A href="ProcurementForm.asp">Submit Procurement Request</A></TD></TR>
	<TR class=SubMenuItem<%=Check("203")%>><TD><A href="HardwareList.asp">Hardware Catalogue</A></TD></TR>
<%End If%>
<TR class=MenuItem<%=Check("300")%>><TD><A href="Training.asp">Training</A></TD></TR>
<%If Left(PageID, 1) = "3" Then%>
	<TR class=SubMenuItem<%=Check("301")%>><TD><A href="TrainManuals.asp">Manuals</A></TD></TR>
	<TR class=SubMenuItem<%=Check("302")%>><TD><A href="TrainLiveMeeting.asp">Live Meeting</A></TD></TR>
<%End If%>
<TR class=MenuItem<%=Check("400")%>><TD><A href="PoliciesAndStandards.asp">Policies and Standards</A></TD></TR>
<%If Left(PageID, 1) = "4" Then%>
	<TR class=SubMenuItem<%=Check("401")%>><TD><A href="PolicyAcceptableUsage.asp">Acceptable Usage Policy</A></TD></TR>
	<TR class=SubMenuItem<%=Check("402")%>><TD><A href="PolicyRemoteAccess.asp">Remote Access</A></TD></TR>
	<TR class=SubMenuItem<%=Check("403")%>><TD><A href="PolicySecurity.asp">Security</A></TD></TR>
	<TR class=SubMenuItem<%=Check("404")%>><TD><A href="PolicyProcurement.asp">Procurement</A></TD></TR>
	<TR class=SubMenuItem<%=Check("405")%>><TD><A href="PolicyUserAccounts.asp">User Accounts</A></TD></TR>
<%End If%>

<TR class=MenuItem<%=Check("700")%>><TD><A href="Projects.asp">Projects</A></TD></TR>

<TR class=MenuItem<%=Check("800")%>><TD><A href="Forms.asp">Forms</A></TD></TR>
<%If Left(PageID, 1) = "8" Then%>
	<TR class=SubMenuItem<%=Check("801")%>><TD><A href="FormITWorkRequest.asp">IT Work Request Form</A></TD></TR>
	<TR class=SubMenuItem<%=Check("802")%>><TD><A href="FormPackageRequest.asp">Application Request Form</A></TD></TR>
	<TR class=SubMenuItem<%=Check("803")%>><TD><A href="FormProjectInitiation.asp">Project Initiation Request</A></TD></TR>
<%End If%>

<TR class=MenuItem<%=Check("500")%>><TD><A href="SelfService.asp">Self Service</A></TD></TR>
<%If Left(PageID, 1) = "5" Then%>
	<TR class=SubMenuItem<%=Check("501")%>><TD><A href="UpdateDetails.asp">Update My Details</A></TD></TR>
	<TR class=SubMenuItem<%=Check("502")%>><TD><A href="LogServiceCall.asp">Log a Service Call</A></TD></TR>
<%
	End If
End Sub
%>