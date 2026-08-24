<!--#include file="Connections/WolLan.asp" -->
<%
  Dim rsUsers
  Set rsUsers = Server.CreateObject("ADODB.Recordset")
  rsUsers.Open "Attending", objConn, adOpenForwardOnly, adLockOptimistic, adCmdTable

	  rsUsers.AddNew
	  rsUsers("Alias") = Request.Form("Alias")
	  rsUsers("Email") = Request.Form("Email")
	  rsUsers("Games") = Request.Form("Games")
	  rsUsers.Update                                               ' update the database

  Response.Redirect "fra_view.asp"
%>