<!--#include file="Database.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Create a LAN</title>
</head>

<body>
<% 
If Request("LAN") = "FALSE" then
	Response.Write "<H3>A LAN already exists on that day!</H3>"
End If
 %>

<form action="addLAN.asp" method="post">

<table>
<tr>
	<td>LAN Date (must be in format 12-Oct-2002):</td>
	<td><input type="text" name="LANDate"></td>
</tr>
<tr>
	<td>Cost:</td>
	<td><input type="text" name="Cost"></td>
</tr>
<tr>
	<td>Max People:</td>
	<td><input type="text" name="MaxPPL"></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td><input type="submit" name="Submit" value="Submit"></td>
</tr>
</table>



</form>
</body>
</html>
