<!--#include file="Database.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>View Attendances</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.standard {
	font-family: Geneva, Arial, Helvetica, san-serif;
	font-size: 12px;
	font-style: normal;
	line-height: normal;
	font-weight: normal;
	font-variant: normal;
	text-transform: none;
	color: #333333;
	text-decoration: none;
}
.content {
	font-family: Geneva, Arial, Helvetica, san-serif;
	font-size: 12px;
	font-style: normal;
	line-height: normal;
	font-weight: normal;
	font-variant: normal;
	text-transform: none;
	color: #000000;
	text-decoration: none;
}
a {
	font-family: Geneva, Arial, Helvetica, san-serif;
	font-size: 12px;
	font-style: normal;
	line-height: normal;
	font-weight: normal;
	font-variant: normal;
	text-transform: none;
	color: #333333;
	text-decoration: none;
}
-->
</style>
</HEAD>
<BODY bgcolor="#666666">

<table border=1>
<tr>
<th><a href=ViewMembers.asp?FilterBy=MemberID>Member ID</a></th>
<th><a href=ViewMembers.asp?FilterBy=Alias>Alias</a></th>
<th><a href=ViewMembers.asp?FilterBy=LastName>Last Name</a></th>
<th><a href=ViewMembers.asp?FilterBy=FirstName>First Name</a></th>
<th>E-Mail Address</th>
</tr>
<% 
  Dim rsView
If Request("FilterBy") = "MemberID" Then
	strSQL = "SELECT * FROM Users ORDER BY UserID;"
End If
If Request("FilterBy") = "Alias" Then
	strSQL = "SELECT * FROM Users ORDER BY UserName;"
End If
If Request("FilterBy") = "FirstName" Then
	strSQL = "SELECT * FROM Users ORDER BY FirstName;"
End If
If Request("FilterBy") = "LastName" Then
	strSQL = "SELECT * FROM Users ORDER BY LastName;"
End If
If Request("FilterBy") = "EMail" Then
	strSQL = "SELECT * FROM Users ORDER BY EMailAddress DESC;"
End If
If Request("FilterBy") = "" Then
	strSQL = "SELECT * FROM Users"
End If

	
	Set rsView = Server.CreateObject("ADODB.Recordset")
	rsView.Open strSQL, objConn
Dim Count  	
  Do While Not rsView.EOF
	Response.Write "<tr align=center>"
	Response.Write "<td>" & rsView("UserID") & "</td>"
	Response.Write "<td>" & rsView("Username") & "</td>"
	Response.Write "<td>" & rsView("LastName") & "</td>"
	Response.Write "<td>" & rsView("Firstname") & "</td>"
	Response.Write "<td><a href=mailto:" & rsView("EMailAddress") & ">" & rsView("EMailAddress") & "</a></td>"
	Response.Write "</tr>"
  Count = Count + 1
  rsView.MoveNext
  Loop
  rsView.close
  Set rsView = Nothing

 %>
	</table>

Number of Members: <%= count %>

</body>
</html>
