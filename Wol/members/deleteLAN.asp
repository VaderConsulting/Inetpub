<!--#include file="Database.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Delete a LAN</title>
</head>

<body>
<table border=1>
<tr>
<th>LAN ID</th>
<th>LAN Date</th>
<th>Cost</th>
<th>Max PPL</th>
<th>Delete</th>
</tr>
<% 

  Dim rsLANSearch 
  strSQL = "SELECT * FROM LANs "
  Set rsLANSearch = Server.CreateObject("ADODB.Recordset")
  rsLANSearch.Open strSQL, objConn
  
  Do While Not rsLANSearch.EOF

Response.Write "<tr align=center>"
	Response.Write "<td>" & rsLANSearch("LanID") & "</td>"
	Response.Write "<td>" & FormatDateTime(rsLANSearch("LANDate"),1) & "</td>"
	Response.Write "<td>$" & rsLANSearch("Cost") & "</td>"
	Response.Write "<td>" & rsLANSearch("MaxPPL") & "</td>"
	Response.Write "<td><a href=delLAN.asp?LanID=" & rsLANSearch("LanID") & ">Delete</a></td>"
Response.Write "</tr>"


  rsLANSearch.MoveNext
  Loop
  rsLANSearch.close
  Set rsLANSearch = Nothing

 %>
	</table>
</body>
</html>
