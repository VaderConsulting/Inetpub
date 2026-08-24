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

		Case "INDEX.ASP":					PageID = "100"
		Case "FORMNEWEMPLOYEE.ASP":			PageID = "101"
		Case "FORMNEWNETWORKACCESS.ASP":	PageID = "102"
		Case "FORMEXTENDACCESS.ASP":		PageID = "103"
		Case "FORMELLIPSEVARYUSER.ASP":		PageID = "104"
		Case "FORMNETWORKVARYUSER.ASP":		PageID = "105"
		Case "FORMTRANSFEREMPLOYEE.ASP":	PageID = "106"
		Case "FORMTERMINATINGEMPLOYEE.ASP":	PageID = "107"
		Case "FORMCHANGEOFADDRESS.ASP":		PageID = "108"
		
		Case "SITEMAP.ASP":					PageID = "600"
		Case "CONTACTUS.ASP":				PageID = "601"
		Case "INTRANETDISCLAIMER.ASP":		PageID = "602"
		Case "INTRANETPRIVACYSTATEMENT.ASP":	PageID = "603"

		Case Else: Response.write TheURL
	End Select
%>
<TR class=MenuItem<%=Check("100")%>><TD><A href="index.asp">Forms</A></TD></TR>
<%If Left(PageID, 1) = "1" Then%>
	<TR class=SubMenuItem<%=Check("101")%>><TD><A href="FormNewEmployee.asp">New Employee Access</A></TD></TR>
	<TR class=SubMenuItem<%=Check("102")%>><TD><A href="FormNewNetworkAccess.asp">New Network Access</A></TD></TR>
	<TR class=SubMenuItem<%=Check("103")%>><TD><A href="FormExtendAccess.asp">Extend User Access</A></TD></TR>
	<TR class=SubMenuItem<%=Check("104")%>><TD><A href="FormEllipseVaryUser.asp">Ellipse Vary User</A></TD></TR>
	<TR class=SubMenuItem<%=Check("105")%>><TD><A href="FormNetworkVaryUser.asp">Network Vary User</A></TD></TR>
	<TR class=SubMenuItem<%=Check("106")%>><TD><A href="FormTransferEmployee.asp">Transfer Employee</A></TD></TR>
	<TR class=SubMenuItem<%=Check("107")%>><TD><A href="FormTerminatingEmployee.asp">Terminate Employee Access</A></TD></TR>
	<TR class=SubMenuItem<%=Check("108")%>><TD><A href="FormChangeOfAddress.asp">Change of Address</A></TD></TR>
<%End If
End Sub
%>