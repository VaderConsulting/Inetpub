<%
'define the variables
Dim PageID

' Define URL constants for Sharepoint
Const NewEmployeeAccessURL = "http://oxnet.oxiana.com.au/hr/Forms/_layouts/FormServer.aspx?XsnLocation=http://oxnet.oxiana.com.au/hr/Forms/New%20Employee%20Request/Forms/template.xsn&SaveLocation=http%3A%2F%2Foxnet%2Eoxiana%2Ecom%2Eau%2Fhr%2FForms%2FNew%20Employee%20Request&Source=http%3A%2F%2Foxnet%2Eoxiana%2Ecom%2Eau%2Fhr%2FForms&DefaultItemOpen=1"
Const TransferEmployeeURL  = "http://oxnet.oxiana.com.au/hr/Forms/_layouts/FormServer.aspx?XsnLocation=http://oxnet.oxiana.com.au/hr/Forms/Transferring%20Employee/Forms/template.xsn&SaveLocation=http%3A%2F%2Foxnet%2Eoxiana%2Ecom%2Eau%2Fhr%2FForms%2FTransferring%20Employee&Source=http%3A%2F%2Foxnet%2Eoxiana%2Ecom%2Eau%2Fhr%2FForms&DefaultItemOpen=1"
Const ChangeOfAddressURL   = "http://oxnet.oxiana.com.au/hr/Forms/_layouts/FormServer.aspx?XsnLocation=http://oxnet.oxiana.com.au/hr/Forms/Change%20employee%20addresscontact%20details/Forms/template.xsn&SaveLocation=http%3A%2F%2Foxnet%2Eoxiana%2Ecom%2Eau%2Fhr%2FForms%2FChange%20employee%20addresscontact%20details&Source=http%3A%2F%2Foxnet%2Eoxiana%2Ecom%2Eau%2Fhr%2FForms&DefaultItemOpen=1"

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

		Case "INDEX.ASP":					PageID = "100"
		Case "FORMNEWEMPLOYEE.ASP":			PageID = "101"
		Case "FORMNEWNETWORKACCESS.ASP":	PageID = "102"
		Case "FORMEXTENDACCESS.ASP":		PageID = "103"
		Case "FORMELLIPSEVARYUSER.ASP":		PageID = "104"
		Case "FORMNETWORKVARYUSER.ASP":		PageID = "105"
		Case "FORMTRANSFEREMPLOYEE.ASP":	PageID = "106"
		Case "FORMTERMINATINGEMPLOYEE.ASP":	PageID = "107"
		Case "FORMCHANGEOFADDRESS.ASP":		PageID = "108"
		Case "FORMROSTERCHANGE.ASP":		PageID = "109"
		Case "FORMCREWCHANGE.ASP":			PageID = "110"
		
		Case "SITEMAP.ASP":					PageID = "600"
		Case "CONTACTUS.ASP":				PageID = "601"
		Case "INTRANETDISCLAIMER.ASP":		PageID = "602"
		Case "INTRANETPRIVACYSTATEMENT.ASP":	PageID = "603"

		Case Else: Response.write TheURL
	End Select
%>
<tr class=MenuItem<%=Check("100")%>><td><a href="index.asp">Forms</a></td></tr>
<%If Left(PageID, 1) = "1" Then%>
	<tr class=SubMenuItem<%=Check("101")%>><td><a href="<%=NewEmployeeAccessURL%>">New Employee Access</a></td></tr>
	<tr class=SubMenuItem<%=Check("102")%>><td><a href="FormNewNetworkAccess.asp">New Network Access</a></td></tr>
	<tr class=SubMenuItem<%=Check("103")%>><td><a href="FormExtendAccess.asp">Extend User Access</a></td></tr>
	<tr class=SubMenuItem<%=Check("104")%>><td><a href="FormEllipseVaryUser.asp">Ellipse Vary User</a></td></tr>
	<tr class=SubMenuItem<%=Check("105")%>><td><a href="FormNetworkVaryUser.asp">Network Vary User</a></td></tr>
	<tr class=SubMenuItem<%=Check("106")%>><td><a href="<%=TransferEmployeeURL%>">Transfer Employee</a></td></tr>
	<tr class=SubMenuItem<%=Check("107")%>><td><a href="FormTerminatingEmployee.asp">Terminate Employee Access</a></td></tr>
	<tr class=SubMenuItem<%=Check("108")%>><td><a href="<%=ChangeOfAddressURL%>">Change of Address</a></td></tr>
	<tr class=SubMenuItem<%=Check("109")%>><td><a href="FormRosterChange.asp">Roster Change</a></td></tr>
	<tr class=SubMenuItem<%=Check("110")%>><td><a href="FormCrewChange.asp">Position Title, Reporting Line or Crew Change</a></td></tr>
<%End If
End Sub
%>