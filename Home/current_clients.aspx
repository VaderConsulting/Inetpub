<!--#include virtual="includes/dataaccess.inc"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<%

Dim i, j, Person, Cols
Dim Update, ClientID, Status

SQL = "EXEC sp_CurrentClients"
adoRS.Open (SQL, adoConn, 1)
%>

<html>
<head>
    <link rel='stylesheet' type='text/css' href='style.css'>
    <title>Empired Current Clients</title>
</head>

<body onload="">

<script language="vbscript">
  
  Function SetStatus(ClientID, ClientStatus)
    Client.ClientID.Value = ClientID
    
    '' // Get a value (0 or 1) according to the checkbox
    ClientStatus = Abs(CInt(ClientStatus))
    Client.Status.Value = ClientStatus
    Client.Update.value = "1"
    Client.submit
  End Function
  
</script>

    <div align='center' class='Body'>
    Empired Current Clients<br>
    (Total = <%=adoRS.RecordCount%>)
    </div>
    <br>
    <span id="Message"></span>
    <div align='left' class='Body'>
<form name="Client" id="Client" action="setCurrent.asp" method="post">
<table align="center" border='0' cellborder='0' cellpadding='0' cellspacing='0'>
  <tr>
<%Cols = 4  
  For i = 1 to Cols%>
    <td class="TableCol1" width="120">Client Name</td>
    <td width="30">&nbsp;</td>
<%Next%>
  </tr>
  <tr>
<%i = 0
Do Until adoRS.eof%>
    <td><%=adoRS("ClientName").Value%></td>
    <td width="30">&nbsp;</td>
<%  adoRS.MoveNext
    i = i + 1
    j = j + 1
    If i = Cols then 
      Response.Write ("</tr>" & vbcrlf)
      i = 0
    End If
Loop
%>
<input type="hidden" name="ClientID" id="ClientID">
<input type="hidden" name="Status" id="Status">
<input type="hidden" name="Update" id="Update" value="0">
</form>
</div>
</body>
</html>
