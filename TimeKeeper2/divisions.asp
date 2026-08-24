<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
	<title>Divisions</title>
</head>

<body>
<%
    Dim adoConn, adoRS
    Dim SQL
    Set adoConn = Server.CreateObject("ADODB.Connection")
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    SQL = "SELECT Name FROM tblDivision ORDER BY Name"
    adoConn.Open Session("DSN")
    adoRS.Open SQL, adoConn
%>

<script language='vbscript'>
    
    Sub EnableDeleteButton
        DeleteButton.InnerHTML = "<input type='button' value='Delete' style='width: 65px;' onclick='DeleteDivision'>"
    End Sub
    
    Sub EnableAddButton
        If frmInput.DivisionName.Value <> "" Then
            AddButton.InnerHTML = "<input type='button' value='Add' style='width: 65px;' onclick='AddDivision'>"
        Else
            AddButton.InnerHTML = "<input type='button' disabled value='Add' style='width: 65px;' onclick='AddDivision'>"
        End If
    End Sub
    
    Sub DeleteDivision
        frmInput.Operation.Value = "DELETE"
        frmInput.DivisionName.Value = frmInput.cmbDivision.Value
        frmInput.submit
    End Sub
    
    Sub AddDivision
        frmInput.Operation.Value = "INSERT"
        frmInput.submit
    End Sub
    
</script>
    <br>
    <form action='processsql.asp' method='post' name='frmInput' id='frmInput'>
    <table border='0' cellspacing='0' cellpadding='0' align='center'>
        <tr>
            <td colspan='3' align='center'>Existing Divisions</td>
        </tr>
        <tr>
            <td colspan='3' width='200'>
                <table border='0' cellspacing='0' cellpadding='0' align='center' width='100%'>
                    <tr>
                        <td colspan='1'>
                            <select name='cmbDivision' id='cmbDivision' size='10' onclick='EnableDeleteButton' style='width: 262px;'>
<%
    Do Until adoRS.EOF
%>
                                <option value='<%=adoRS("Name")%>'> <%=adoRS("Name")%>
<%
        adoRS.MoveNext
    Loop
%>
                            </select>
                        </td>
                        <td valign='bottom' align='right'><span id='DeleteButton'><input type='button' disabled value='Delete' style='width: 65px;' onclick='DeleteDivision'></span></td>
                    </tr>
                    
                </table>
            </td>
        </tr>
        <tr>
            <td>New Division:</td>
            <td width='100px'><input type="text" name="DivisionName" id="DivisionName" onKeyUp="EnableAddButton"></td>
            <td align='right'><span id='AddButton'><input type='button' disabled name='Add' id='Add' value='Add' style='width: 65px;' onclick='AddDivision'></span></td>
        </tr>
    </table>
    <input type='hidden' name='Operation' id='Operation' value=''>
    <input type='hidden' name='Tablename' id='Tablename' value='tblDivision'>
    </form>
<%
    adoRS.Close
    adoConn.Close
    Set adoRS=Nothing
    Set adoConn=Nothing
%>

</body>
</html>
