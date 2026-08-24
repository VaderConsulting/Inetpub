<!--#include file="Database.asp"-->
<% 

  Dim rsAttending
  Set rsAttending = Server.CreateObject("ADODB.Recordset")
  rsAttending.Open "Attendings", objConn, adOpenForwardOnly, adLockOptimistic, adCmdTable
  rsAttending.Filter = "LanID = " & Request("LanID")
  If Not rsAttending.EOF then
	  Do While Not rsAttending.EOF
		rsAttending.Delete
		rsAttending.MoveNext
	  Loop
  End If

rsAttending.Close
Set rsAttending = Nothing

  Dim rsLANDel
  Set rsLANDel = Server.CreateObject("ADODB.Recordset")
  rsLANDel.Open "LANs", objConn, adOpenForwardOnly, adLockOptimistic, adCmdTable
  	rsLANDel.Filter = "LanID = " & Request("LanID")
	rsLANDel.Delete


rsLANDel.Close
Set rsLANDel = Nothing



 %>