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
    
    Sub EnableButtons
        DeleteButton.InnerHTML = "<input type='button' value='Delete' style='width: 65px;' onclick='DeleteTask'>"
        EditButton.InnerHTML = "<input type='button' value='Edit' style='width: 65px;' onclick='EditTask'>"
    End Sub
    
    Sub EnableAddButton
        If frmTask.NewTaskName.Value <> "" Then
            AddButton.InnerHTML = "<input type='button' value='Add' style='width: 65px;' onclick='AddTask'>"
        Else
            AddButton.InnerHTML = "<input type='button' disabled value='Add' style='width: 65px;' onclick='AddTask'>"
        End If
    End Sub
    
    Sub AddTask
        frmTask.Submit
    End Sub
    
    Sub DeleteTask
        frmTask.Operation.Value = "DELETE"
        frmTask.Submit
    End Sub
    
    Sub EditTask
        frmTask.Action = "edittask.asp"
        frmTask.Submit
    End Sub
    
    Sub OKProject
        frmProject.Submit
    End Sub
    
    Sub OKTask
    End Sub
</script>

<%
    Dim adoConn, adoRS, ClientID
    Dim SQL
    Set adoConn = Server.CreateObject("ADODB.Connection")
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    
    ClientID = Request("ClientID")
    ProjectID = Request("ProjectID")
    
    'Response.Write "Client  ID: " & ClientID & "<br>"
    'Response.Write "Project ID: " & ProjectID & "<br>"
    
    ' Get Clients
    SQL = "SELECT ClientID, ClientName, ClientCode FROM tblClient ORDER BY Clientname"
    adoConn.Open Session("DSN")
    adoRS.Open SQL, adoConn
%>
    <br>
    <form action='tasks.asp' method='post' name='frmClient' id='frmClient'>
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
    If CStr(ClientID) <> "" Then
        'Get Projects
        SQL = "SELECT ProjectName, ProjectCode, ProjectID FROM tblProject WHERE ClientID = " & ClientID & " ORDER BY Projectname"
        adoRS.Open SQL, adoConn
%>
        <br>
        <form action='tasks.asp' method='post' name='frmProject' id='frmProject'>
        <table border='0' cellspacing='0' cellpadding='0' align='center'>
            <tr>
                <td colspan='1' >Choose a Project</td>
            </tr>
            <tr>
                <td colspan='1' width='200'>
                    <select name='ProjectID' id='ProjectID' size='1' style='width: 262px;'>
<%
        Do Until adoRS.EOF
%>
                        <option value='<%=adoRS("ProjectID")%>'<%=adoRS("ProjectID")%>'<%If ProjectID = cstr(adoRS("ProjectID")) Then %> selected<%End If%>><%=adoRS("ProjectName")%> (<%=adoRS("ProjectCode")%>)
<%
            adoRS.MoveNext
        Loop
%>
                    </select>
                </td>
                <td valign='bottom' align='right'><span id='ProjectButton'><input type='button' value='OK' style='width: 65px;' onclick='OKProject'></span></td>
            </tr>
        </table>
        <input type='hidden' name='ClientID' id='ClientID' value='<%=ClientID%>'>
        <input type='hidden' name='Operation' id='Operation' value=''>
        <input type='hidden' name='Tablename' id='Tablename' value='tblTask'>
        </form>
<%
        adoRS.Close
    End If
    If ProjectID <> "" Then
        'Get Tasks
        SQL = "SELECT TaskID, Description FROM tblTask WHERE ProjectID = " & ProjectID & " ORDER BY Description"
        adoRS.Open SQL, adoConn
%>
        <br>
        <span id = 'taskdestination'><form action='processsql.asp' method='post' name='frmTask' id='frmTask'></span>
        <table border='0' cellspacing='0' cellpadding='0' align='center'>
            <tr>
                <td colspan='3' >Task</td>
            </tr>
            <tr>
                <td colspan='3' width='200'>
                    <table border='0' cellspacing='0' cellpadding='0' align='center' width='100%'>
                        <tr>
                            <td colspan='1'>
                                <select name='TaskID' id='TaskID' size='5' style='width: 262px;' onclick='EnableButtons'>
<%
        Do Until adoRS.EOF
%>
                                    <option value='<%=adoRS("TaskID")%>'><%=adoRS("Description")%>
<%
            adoRS.MoveNext
        Loop
%>
                                </select>
                            </td>
                            <td valign='bottom' align='right'>
                                <span id='EditButton'><input type='button' disabled value='Edit' style='width: 65px;' onclick='EditTask'></span><br>
                                <span id='DeleteButton'><input type='button' disabled value='Delete' style='width: 65px;' onclick='DeleteTask'></span>
                            </td>
                        </tr>
                        
                    </table>
                </td>
            </tr>
            <tr>
                <td>New task:</td>
                <td width='100px'><input type="text" name="NewTaskName" id="NewTaskName" onKeyUp="EnableAddButton"></td>
                <td align='right'><span id='AddButton'><input type='button' disabled name='Add' id='Add' value='Add' style='width: 65px;' onclick='AddTask'></span></td>
            </tr>
        </table>
        <input type='hidden' name='ClientID' id='ClientID' value='<%=ClientID%>'>
        <input type='hidden' name='ProjectID' id='ClientID' value='<%=ProjectID%>'>
        <input type='hidden' name='Operation' id='Operation' value='INSERT'>
        <input type='hidden' name='Tablename' id='Tablename' value='tblTask'>
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
