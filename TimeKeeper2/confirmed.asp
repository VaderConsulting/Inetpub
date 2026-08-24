<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
	<title>Confirmation</title>
</head>
<%
    Dim adoConn, adoRS
    Dim SQL, Tablename
    Set adoConn = Server.CreateObject("ADODB.Connection")
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    SQL = Request("SQL")
    Tablename = Request("TableName")
    adoConn.Open Session("DSN")
    'Response.Write SQL & vbcrlf
    adoConn.Execute SQL
%>
<body>
<br>
<div align='center'>
<%=Tablename%> table has been updated.
<%
Response.Write "<form action='clients.asp' name='frmReturn'>" & vbcrlf
Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
Response.Write "</form>" & vbcrlf
Response.Write "<script language='vbscript'>" & vbcrlf
Response.Write "Sub Window_OnLoad" & vbcrlf
'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
Response.Write "    frmReturn.Submit" & vbcrlf
Response.Write "End Sub" & vbcrlf
Response.Write "</script>" & vbcrlf
%>
</div>
</body>
</html>
