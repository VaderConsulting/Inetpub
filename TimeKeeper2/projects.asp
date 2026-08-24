<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
	<title>Projects</title>
</head>

<body>

<script language='vbscript'>
    
    Sub EnableOKButton
        OKButton.InnerHTML = "<input type='button' value='OK' style='width: 65px;' onclick='OKButton'>"
    End Sub
    
    Sub OKButton
        frmClient.Submit
    End Sub
    
    Sub EnableDeleteButton
        DeleteButton.InnerHTML = "<input type='button' value='Delete' style='width: 65px;' onclick='DeleteProject'>"
    End Sub
    
    Sub EnableAddButton
        If frmProject.ProjectName.Value <> "" Then
            AddButton.InnerHTML = "<input type='button' value='Add' style='width: 65px;' onclick='AddProject'>"
        Else
            AddButton.InnerHTML = "<input type='button' disabled value='Add' style='width: 65px;' onclick='AddProject'>"
        End If
    End Sub
    
    Sub AddProject
        frmProject.Operation.Value = "INSERT"
        frmProject.Submit
    End Sub
    
    Sub DeleteProject
        frmProject.Operation.Value = "DELETE"
        frmProject.ProjectName.Value = frmProject.cmbProject.Value
        frmProject.Submit
    End Sub
</script>

<%
    Dim adoConn, adoRS, ClientID
    Dim SQL
    Set adoConn = Server.CreateObject("ADODB.Connection")
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    
    ClientID = Request("ClientID")
    
    ' Get Clients
    SQL = "SELECT ClientID, ClientName, ClientCode FROM tblClient ORDER BY Clientname"
    adoConn.Open Session("DSN")
    adoRS.Open SQL, adoConn
%>
    <br>
    <form action='projects.asp' method='post' name='frmClient' id='frmClient'>
    <table border='0' cellspacing='0' cellpadding='0' align='center'>
        <tr>
            <td colspan='2'>Choose a client</td>
        </tr>
        <tr>
            <td>
                <select name='ClientID' id='ClientID' size='1' style='width: 262px;'>
<%
    Do Until adoRS.EOF
%>
                    <option value='<%=adoRS("ClientID")%>'<%If ClientID = cstr(adoRS("ClientID")) Then %> selected<%End If%>><%=adoRS("ClientName")%> (<%=adoRS("ClientCode")%>)
<%
        adoRS.MoveNext
    Loop
%>
                </select>
            </td>
            <td valign='bottom' align='right'><span id='OKButton'><input type='button' value='OK' style='width: 65px;' onclick='OKButton'></span></td>
        </tr>
    </table>
    </form>
<%  
    adoRS.Close
    If ClientID <> "" and ClientID <> 0 Then
        'Get Projects
        SQL = "SELECT ProjectName, ProjectCode FROM tblProject WHERE ClientID = " & ClientID & " ORDER BY Projectname"
        adoRS.Open SQL, adoConn
%>
        <br>
        <form action='processsql.asp' method='post' name='frmProject' id='frmProject'>
        <table border='0' cellspacing='0' cellpadding='0' align='center'>
            <tr>
                <td colspan='3' >Projectname (Project Code)</td>
            </tr>
            <tr>
                <td colspan='3' width='200'>
                    <table border='0' cellspacing='0' cellpadding='0' align='center' width='100%'>
                        <tr>
                            <td colspan='1'>
                                <select name='cmbProject' id='cmbProject' size='10' onclick='EnableDeleteButton' style='width: 262px;'>
<%
        Do Until adoRS.EOF
%>
                                    <option value='<%=adoRS("ProjectName")%>'<%If ProjectName = adoRS("ProjectName") Then %> selected<%End If%>><%=adoRS("ProjectName")%> (<%=adoRS("ProjectCode")%>)
<%
            adoRS.MoveNext
        Loop
%>
                                </select>
                            </td>
                            <td valign='bottom' align='right'><span id='DeleteButton'><input type='button' disabled value='Delete' style='width: 65px;' onclick='DeleteProject'></span></td>
                        </tr>
                        
                    </table>
                </td>
            </tr>
            <tr>
                <td>New project:</td>
                <td width='100px'><input type="text" name="ProjectName" id="ProjectName" onKeyUp="EnableAddButton"></td>
                <td align='right'><span id='AddButton'><input type='button' disabled name='Add' id='Add' value='Add' style='width: 65px;' onclick='AddProject'></span></td>
            </tr>
        </table>
        <input type='hidden' name='Operation' id='Operation' value=''>
        <input type='hidden' name='ClientID' id='ClientID' value='<%=ClientID%>'>
        <input type='hidden' name='Tablename' id='Tablename' value='tblProject'>
        </form>
<%
        adoRS.Close
    End If
    adoConn.Close
    Set adoRS=Nothing
    Set adoConn=Nothing
%>
</body>
</html>
