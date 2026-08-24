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
<th>Member ID</th>
<th>Alias</th>
</tr>

<% 
  Dim rsView, Count
	strSQL = "SELECT * FROM View_Attending_Query WHERE LanID = " & Request("LanID")
	Set rsView = Server.CreateObject("ADODB.Recordset")
	rsView.Open strSQL, objConn

  	  	Count = 0
  	
  Do While Not rsView.EOF
	Response.Write "<tr align=center>"
	Response.Write "<td>" & rsView("UserID") & "</td>"
	Response.Write "<td>" & rsView("Username") & "</td>"
	Response.Write "</tr>"

  rsView.MoveNext
  
			Count = Count + 1
			
  Loop
		
	
  rsView.close
  Set rsView = Nothing

%>
	</table>
<%
		Response.Write "Number of People Signed Up = " & Count
%>
</body>
</html>