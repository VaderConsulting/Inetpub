<!--#include virtual="includes/dataaccess.inc"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<%
Dim Update, ClientID, Status

' Get page variables
  Update = Request.Form("Update")
  Status = Request.Form("Status")
  ClientID = Request.Form("ClientID")

%>

<html>
<head>
    <link rel='stylesheet' type='text/css' href='style.css'>
    <META HTTP-EQUIV="REFRESH" CONTENT=60>
    <title>Set Current</title>
</head>

<body onload="Go_Clients()">

<script language="VBScript">
  Function Go_Clients
    window.location = "all_clients.aspx"
  End Function
</script>

    <div align='center' class='Body'>
    </div>
    <br>
    <div align='left' class='Body'>
<%
  If Update = "1" Then
    SQL = "EXEC SP_UpdateClient " & ClientID & "," & Status
    adoConn.Execute (SQL)
    '' // Response.Write SQL
  End If
%>
</body>
</html>
