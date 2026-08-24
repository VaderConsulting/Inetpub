<!--#include file="Database.asp"-->
<% 

  Dim rsLANSearch 
  strSQL = "SELECT * FROM LANs "
  Set rsLANSearch = Server.CreateObject("ADODB.Recordset")
  rsLANSearch.Open strSQL, objConn
  
  Do While Not rsLANSearch.EOF
  
  	If FormatDateTime(Request.Form("LANDate"),0) = FormatDateTime(rsLANSearch("LANDate"),0) Then
		Response.Redirect "createLAN.asp?LAN=False"
	End If
  rsLANSearch.MoveNext
  Loop
  rsLANSearch.close
  Set rsLANSearch = Nothing




  Dim rsLANs
  Set rsLANs = Server.CreateObject("ADODB.Recordset")
  rsLANs.Open "LANs", objConn, adOpenForwardOnly, adLockOptimistic, adCmdTable

rsLANs.AddNew

rsLANs("LANDate") = Request.Form("LANDate")
rsLANs("Cost") = Request.Form("Cost")
rsLANs("MaxPPL") = Request.Form("MaxPPL")
rsLANs.Update

rsLANs.Close
Set rsLANs = Nothing

Response.Redirect "menu_for_registered_users.asp?createLAN=Successful"

 %>
