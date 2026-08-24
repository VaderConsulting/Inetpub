<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
	<title>Clients</title>
</head>

<body>
<%
    Dim adoConn, adoRS
    Dim SQL
    Set adoConn = Server.CreateObject("ADODB.Connection")
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT ClientName, ClientCode FROM tblClient ORDER BY Clientname"
    adoConn.Open Session("DSN")
    adoRS.Open SQL, adoConn
%>

<script language='vbscript'>
    
    Sub EnableDeleteButton
        DeleteButton.InnerHTML = "<input type='button' value='Delete' style='width: 65px;' onclick='DeleteClient'>"
    End Sub
    
    Sub EnableAddButton
        If frmInput.ClientName.Value <> "" Then
            AddButton.InnerHTML = "<input type='button' value='Add' style='width: 65px;' onclick='AddClient'>"
        Else
            AddButton.InnerHTML = "<input type='button' disabled value='Add' style='width: 65px;' onclick='AddClient'>"
        End If
    End Sub
    
    Sub DeleteClient
        frmInput.Operation.Value = "DELETE"
        frmInput.ClientName.Value = frmInput.cmbClient.Value
        frmInput.submit
    End Sub
    
    Sub AddClient
        frmInput.Operation.Value = "INSERT"
        frmInput.submit
    End Sub
    
</script>
    <br>
    <form action='processsql.asp' method='post' name='frmInput' id='frmInput'>
    <table border='0' cellspacing='0' cellpadding='0' align='center'>
        <tr>
            <td colspan='3' align='center'>Existing Clients</td>
        </tr>
        <tr>
            <td colspan='3' >Clientname (Client Code)</td>
        </tr>
        <tr>
            <td colspan='3' width='200'>
                <table border='0' cellspacing='0' cellpadding='0' align='center' width='100%'>
                    <tr>
                        <td colspan='1'>
                            <select name='cmbClient' id='cmbClient' size='10' onclick='EnableDeleteButton' style='width: 262px;'>
<%
    Do Until adoRS.EOF
%>
                                <option value='<%=adoRS("ClientName")%>'> <%=adoRS("ClientName")%> (<%=adoRS("ClientCode")%>)
<%
        adoRS.MoveNext
    Loop
%>
                            </select>
                        </td>
                        <td valign='bottom' align='right'><span id='DeleteButton'><input type='button' disabled value='Delete' style='width: 65px;' onclick='DeleteClient'></span></td>
                    </tr>
                    
                </table>
            </td>
        </tr>
        <tr>
            <td>New clientname:</td>
            <td width='100px'><input type="text" name="ClientName" id="ClientName" onKeyUp="EnableAddButton"></td>
            <td align='right'><span id='AddButton'><input type='button' disabled name='Add' id='Add' value='Add' style='width: 65px;' onclick='AddClient'></span></td>
        </tr>
    </table>
    <input type='hidden' name='Operation' id='Operation' value=''>
    <input type='hidden' name='Tablename' id='Tablename' value='tblClient'>
    </form>
<%
    adoRS.Close
    adoConn.Close
    Set adoRS=Nothing
    Set adoConn=Nothing
%>

</body>
</html>
