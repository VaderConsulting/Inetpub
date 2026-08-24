<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
	<title>Users</title>
</head>

<body>
<%
    Dim adoConn, adoRS
    Dim SQL
    Set adoConn = Server.CreateObject("ADODB.Connection")
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT Username, FName, LName FROM tblUsers ORDER BY Username"
    adoConn.Open Session("DSN")
    adoRS.Open SQL, adoConn
%>

<script language='vbscript'>
    
    Sub EnableButtons
        EditButton.InnerHTML = "<input type='button' value='Edit' style='width: 65px;' onclick='EditUser'>"
        DeleteButton.InnerHTML = "<input type='button' value='Delete' style='width: 65px;' onclick='DeleteUser'>"
    End Sub
    
    Sub EnableAddButton
        If frmInput.UserName.Value <> "" Then
            AddButton.InnerHTML = "<input type='button' value='Add' style='width: 65px;' onclick='AddUser'>"
        Else
            AddButton.InnerHTML = "<input type='button' disabled value='Add' style='width: 65px;' onclick='AddUser'>"
        End If
    End Sub
    
    Sub DeleteUser
        frmInput.Operation.Value = "DELETE"
        frmInput.UserName.Value = frmInput.cmbUser.Value
        frmInput.submit
    End Sub
    
    Sub AddUser
        frmInput.Operation.Value = "INSERT"
        frmInput.submit
    End Sub
    
    Sub EditUser
        frmInput.Operation.Value = "UPDATE"
        frmInput.UserName.Value = frmInput.cmbUser.Value
        frmInput.submit
    End Sub
    
</script>
    <br>
    <form action='userinfo.asp' method='post' name='frmInput' id='frmInput'>
    <table border='0' cellspacing='0' cellpadding='0' align='center'>
        <tr>
            <td colspan='3' align='center'>Existing Users</td>
        </tr>
        <tr>
            <td colspan='3' >Username</td>
        </tr>
        <tr>
            <td colspan='3' width='200'>
                <table border='0' cellspacing='0' cellpadding='0' align='center' width='100%'>
                    <tr>
                        <td colspan='1'>
                            <select name='cmbUser' id='cmbUser' size='10' onclick='EnableButtons' style='width: 262px;'>
<%
    Do Until adoRS.EOF
%>
                                <option value='<%=adoRS("UserName")%>'> <%=adoRS("UserName") & " (" & adoRS("FName") & " " & adoRS("LName") & ")"%>
<%
        adoRS.MoveNext
    Loop
%>
                            </select>
                        </td>
                        <td valign='bottom' align='right'>
                            <span id='EditButton'><input type='button' disabled value='Edit' style='width: 65px;' onclick='EditUser'></span><br>
                            <span id='DeleteButton'><input type='button' disabled value='Delete' style='width: 65px;' onclick='DeleteUser'></span>
                        </td>
                    </tr>
                    
                </table>
            </td>
        </tr>
        <tr>
            <td>New Username:</td>
            <td width='100px'><input type="text" name="UserName" id="UserName" onKeyUp="EnableAddButton"></td>
            <td align='right'><span id='AddButton'><input type='button' disabled name='Add' id='Add' value='Add' style='width: 65px;' onclick='AddUser'></span></td>
        </tr>
    </table>
    <input type='hidden' name='Operation' id='Operation' value=''>
    <input type='hidden' name='Tablename' id='Tablename' value='tblUsers'>
    </form>
<%
    adoRS.Close
    adoConn.Close
    Set adoRS=Nothing
    Set adoConn=Nothing
%>

</body>
</html>
