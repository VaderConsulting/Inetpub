<!--#include virtual="includes/dataaccess.inc"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<%
Dim i, j
Dim Person, Returned

' Get page variables
  Returned = Request.QueryString("Returned")
  If Returned = "1" Then
    Person = Request.QueryString("Person")
  End If

%>

<html>
<head>
    <link rel='stylesheet' type='text/css' href='style.css'>
    <title>Set Returned</title>
</head>

<body onload="Go_Locator()">

<script language="VBScript">
  Function Go_Locator
    window.location = "locator.aspx"
  End Function
</script>

    <div align='center' class='Body'>
    </div>
    <br>
    <div align='left' class='Body'>
<%
  ' Update User location
  ' Example: sp_UpdateLocatorTable 1,2,'10:30:00 AM', '11:30:00 AM'
  If Returned = "1" Then
    SQL = "EXEC SP_UpdateLocatorTable " & Person & ",1,'',''"
    adoConn.Execute (SQL)
  End If
%>
</body>
</html>
