<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="Connections/WolLan.asp" -->
<%
Dim Recordset1
Dim i 
Dim b
Set Recordset1 = Server.CreateObject("ADODB.Recordset")
  strSQL = "SELECT * FROM Attending"
  Recordset1.Open strSQL, objConn
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/standard.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>View</title>
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
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable --> 
</head>

<body bgcolor="#666666">
 
      <td bordercolor="#000000" bgcolor="#666666"> <!-- InstanceBeginEditable name="centerContent" --> 
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
<tr>
                       
            <td bgcolor="#333333" class="content" width="50%"><font color="#CCCCCC"><strong><p align="center">Alias</strong></font></p></td>
            <td bgcolor="#333333" class="content" width="50%"><font color="#CCCCCC"><strong><p align="center">Games</strong></font></p></td>
          </tr>
<% 
Do While i < 80 and Not Recordset1.EOF
%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0">
<tr> 
            <td width="50%"><p align="left"><%=(Recordset1("Alias"))%></p></td>
            <td width="50%"><p align="left"><%=(Recordset1("Games"))%></p></td>
          </tr>
        </table>
<% 
Recordset1.MoveNext
i = i + 1
Loop
%>
<%
Do while i =80
%>
     <tr>
            <td bgcolor="#333333" class="content" width="50%">
      <font color="#FFFFFF"><strong><p align="center">Over Booked</strong></font></p></td>
          </tr>

<% 
Do While i =>80 and not recordset1.EOF
%>
        <table width="100%" border="0" cellpadding="2" cellspacing="0">
<tr> 
            <td width="50%"><font color="#FFFFFF"><p align="left"><%=(Recordset1("Alias"))%></p></td>
            <td width="50%"><font color="#FFFFFF"><p align="left"><%=(Recordset1("Games"))%></p></td>
          </tr>
        </table>
<% 
Recordset1.MoveNext
i = i + 1
b = b + 1
Loop
loop

Response.Write "<BR>People Attending: " & (i-b)

Do while i > 80 
   Response.Write "<BR>Over Booked: " & b
i=0
Loop
%>
</body>
</html>
<%
Recordset1.Close
Set Recordset1 = Nothing
%>