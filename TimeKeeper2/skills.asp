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
    SQL = "SELECT Description FROM tblSkills ORDER BY Description"
    adoConn.Open Session("DSN")
    adoRS.Open SQL, adoConn
%>

<script language='vbscript'>
    
    Sub EnableDeleteButton
        DeleteButton.InnerHTML = "<input type='button' value='Delete' style='width: 65px;' onclick='DeleteSkill'>"
    End Sub
    
    Sub EnableAddButton
        If frmSkill.SkillName.Value <> "" Then
            AddButton.InnerHTML = "<input type='button' value='Add' style='width: 65px;' onclick='AddSkill'>"
        Else
            AddButton.InnerHTML = "<input type='button' disabled value='Add' style='width: 65px;' onclick='AddSkill'>"
        End If
    End Sub
    
    Sub AddSkill
        frmSkill.Operation.Value = "INSERT"
        frmSkill.Submit
    End Sub
    
    Sub DeleteSkill
        frmSkill.Operation.Value = "DELETE"
        frmSkill.SkillName.Value = frmSkill.cmbSkill.Value
        frmSkill.Submit
    End Sub
    
</script>
    <br>
    <form action='processsql.asp' method='post' name='frmSkill' id='frmSkill'>
    <table border='0' cellspacing='0' cellpadding='0' align='center'>
        <tr>
            <td colspan='3' align='center'>Existing Skills</td>
        </tr>
        <tr>
            <td colspan='3' >Skill</td>
        </tr>
        <tr>
            <td colspan='3' width='200'>
                <table border='0' cellspacing='0' cellpadding='0' align='center' width='100%'>
                    <tr>
                        <td colspan='1'>
                            <select name='cmbSkill' id='cmbSkill' size='10' onclick='EnableDeleteButton' style='width: 262px;'>
<%
    Do Until adoRS.EOF
%>
                                <option value='<%=adoRS("Description")%>'> <%=adoRS("Description")%>
<%
        adoRS.MoveNext
    Loop
%>
                            </select>
                        </td>
                        <td valign='bottom' align='right'><span id='DeleteButton'><input type='button' disabled value='Delete' style='width: 65px;' onclick='DeleteSkill'></span></td>
                    </tr>
                    
                </table>
            </td>
        </tr>
        <tr>
            <td>New skill:</td>
            <td width='100px'><input type='text' name='SkillName' id='SkillName' onKeyUp='EnableAddButton'></td>
            <td align='right'><span id='AddButton'><input type='button' disabled name='Add' id='Add' value='Add' style='width: 65px;' onclick='AddSkill'></span></td>
        </tr>
    </table>
    <input type='hidden' name='Operation' id='Operation' value=''>
    <input type='hidden' name='Tablename' id='Tablename' value='tblSkills'>
    </form>
<%
    adoRS.Close
    adoConn.Close
    Set adoRS=Nothing
    Set adoConn=Nothing
%>

</body>
</html>
