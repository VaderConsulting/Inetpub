<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!--#include file="./include/db.inc"-->
<!--#include file="./include/global.inc"-->
<%
    response.expires = -1
%>

<html>
<head>
	<title>Add Host</title>
</head>

<body>
<div align="center">
<H2>Add Host</H1>
<hr>
<form name="frmAdd" id="frmAdd" action = "POST" target = "_self">
    <table>
        <tr> 
            <td>Hostname</td>
            <td><input type ="text" size="30" maxlength="30" name = "Hostname" id = "Hostname"></td>
            <td>IP Address</td>
            <td><input type ="text" size ="13" maxlength="15" name = "IP" id = "IP"></td>
        <tr> 
            <td>Route</td>
            <td><input type="text" size="30" name="Route" id="Route" maxlength="255"></td>
            <td>Timeout</td>
            <td><input type="text" size="5" name="Timeout" id="Timeout" maxlength="5"></td>
        </tr>
        <tr>
            <td valign="top">Email Address</td>
            <td colspan="3">
<%              SQL = "SELECT Address FROM tblEmail ORDER BY Address"
                adoConn.open DSN
                adoRS.open SQL, adoConn%>
                <select name="Email" id="Email" size="10" maxlength="50">
<%                  Do until adoRS.eof%>
                        <option type="select"><%=adoRS("Address")%></option>
<%                      adoRS.movenext
                    Loop%>
                </select>
            </td>
        <tr>
            <td>&nbsp;</td>
        </tr>
        <tr>
            <td colspan="4" align="center"><input type ="button" size="10" name = "cmdAdd" id="cmdAdd" value = "Add">
        </tr>
   </table>
</form>
<hr>
</div>
</body>
</html>
