<!-- METADATA TYPE="typelib"
				 FILE="C:\Program Files\Common Files\System\ado\msado15.dll" -->

<%
	Dim objConn
	Set objConn = Server.CreateObject("ADODB.Connection")
	objConn.Open "Provider=Microsoft.Jet.OLEDB.4.0; " & _
					"Data Source= c:\Inetpub\wwwroot\wol\Data\mainData.mdb"

%>
