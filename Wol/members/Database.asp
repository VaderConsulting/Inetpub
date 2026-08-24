<!-- METADATA TYPE="typelib"
				 FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->

<%
	Dim objConn
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; " & _
					"Data Source= C:\inetpub\wwwroot\wol\members\Wol.mdb"

If Session("blnValidUser") = True and Session("UserID") = " " Then
	Dim rsUserIDCheck
	Set  rsUserIDCheck = Server.CreateObject("ADODB.Recordset")
	Dim strSQL
	strSQL = "SELECT UserID FROM Users " & _
			   "WHERE Username = '" & Session("Username") & "';"
	rsUserIDCheck.Open strSQL, objConn
	If rsUserIDCheck.EOF Then
		Session("blnValidUser") = False
	Else
		Session("UserID") = rsStudentIDCheck("UserID")
	End IF
	rsUserIDCheck.Close
	Set rsUserIDCheck = Nothing
	End If
%>