<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
	<title>User Info</title>
</head>
<%
    'Dim ClientName, Operation, ClientCode, ClientID, ProjectID, NewTaskName, SelectedTaskName
    Dim adoConn, adoRS, adoRS2
    Dim SQL
    Set adoConn = Server.CreateObject("ADODB.Connection")
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    Set adoRS2 = Server.CreateObject("ADODB.Recordset")
    'ClientName = Request("ClientName")
    'ClientID = Request("ClientID")
    'ProjectID = Request("ProjectID")
    'ProjectName = Request("ProjectName")
    'NewTaskName = Request("NewTaskName")
    'SelectedTaskName = Request("SelectedTaskName")
    'Skillname = Request("Skillname")
    'DivisionName = Request("DivisionName")
    'Operation = Request("Operation")
    'Tablename = Request("Tablename")
    Username = Request("Username")
    
    adoConn.Open Session("DSN")
%>
<body>
<%
SQL =       "SELECT tblUsers.*, tblDivision.Name AS Division FROM tblUsers LEFT OUTER JOIN "
SQL = SQL & "tblDivision On tblUsers.DivisionID = tblDivision.DivisionID WHERE Username = '" & Username & "'"
adoRS2.Open SQL, adoConn
If Not adoRS2.EOF Then
    FName = adoRS2("FName")
    LName = adoRS2("LName")
    Hours = adoRS2("HoursperWeek")
    Contractor = adoRS2("Contractor")
    Division = adoRS2("Division")
End If
adoRS2.Close
%>
<br>
<form action='userinfo.asp' name='frmEdit'>
<table border='0' cellspacing='0' cellpadding='0' align='center'>
    <tr>
        <td align='center'>Username</td>
        <td align='center'>First name</td>
        <td align='center'>Last name</td>
    </tr>
    <tr>
        <td align='center' class='highlight'><strong><%=Username%></strong</td>
        <td align='center'><input type='text' name='FName' id='FName' value='<%=FName%>' style='width: 145px;' class='highlight'></td>
        <td align='center'><input type='text' name='LName' id='LName' value='<%=LName%>' style='width: 262px;' class='highlight'></td>
    </tr>
    <tr>
        <td colspan='3'>&nbsp;</td
    </tr>
    <tr>
        <td align='center'>Hours per week</td>
        <td align='center'>Contractor</td>
        <td align='center'>Division</td>
    </tr>
    <tr>
        <td align='center'><input type='text' name='Hours' id='Hours' value='<%=Hours%>' style='width: 50px;' class='highlight'></td>
        <td align='center'><input type='checkbox' name='Contractor' id='Contractor' value=''<%If Contractor Then %> checked<%End If%> class='highlight'></td>
        <td align='center'>
            <select name='cmbDivision' id='cmbDivision' size='1' style='width: 262px;'>
<%
                SQL = "SELECT DivisionID, Name FROM tblDivision ORDER BY Name"
                adoRS2.Open SQL, adoConn
                Do Until adoRS2.EOF
%>
                    <option value='<%=adoRS2("DivisionID")%>'<%If adoRS2("Name") = Division Then %> selected<%End If%> class='highlight'><%=adoRS2("Name")%>
<%
                    adoRS2.MoveNext
                Loop
                adoRS2.Close
%>
            </select>
        </td>
    </tr>
    <tr>
        <td colspan='3'>&nbsp;</td
    </tr>
    <tr>
        <td colspan='3'><hr></td
    </tr>
    <tr>
        <td align='center'>Skill List</td>
        <td>&nbsp;</td>
        <td align='center'>Personal Skills</td>
    </tr>
    <tr>
<!-- START -->
<%
        ' Separate skills into two tables:
        ' Skills the user does not have (TableA), and those the user does have (TableB).
        SQL =       "SELECT tblSkills.Description, tblSkills.SkillID, tblUsers.Username "
        SQL = SQL & "FROM tblUsers RIGHT JOIN (tblSkills LEFT JOIN tblSkillList ON tblSkills.SkillID = tblSkillList.SkillID) ON tblUsers.Username = tblSkillList.Username "
        SQL = SQL & "WHERE tblUsers.Username = '" & Username & "' OR isNull(tblUsers.Username, '') = '' " 
        SQL = SQL & "ORDER BY tblSkills.Description"
        Response.Write "<!!- " & vbcrlf & SQL & vbcrlf & "-->"
        adoRS.Open SQL, adoConn
        Do Until adoRS.EOF
            If adoRS("Username") & "" = Username Then
                j = j + 1
                ReDim Preserve TableB_ID(j)
                ReDim Preserve TableB_Description(j)
                TableB_ID(j) = adoRS("SkillID")
                TableB_Description(j) = adoRS("Description")
            Else
                i = i + 1
                ReDim Preserve TableA_ID(i)
                ReDim Preserve TableA_Description(i)
                TableA_ID(i) = adoRS("SkillID")
                TableA_Description(i) = adoRS("Description")
            End If
            adoRS.MoveNext
        Loop
        adoRS.Close
%>
        <td width='300' align='center'>
            <select name='cmbSkillsA' id='cmbSkillsB' size='10' style='width: 262px;'>
<%
                For a = 1 To i
%>
                    <option value='<%=TableA_ID(a)%>'><%=TableA_Description(a)%>
<%
                Next
%>
            </select>
        </td>
        <td align='center'>
            <input type='button' name='AddSkill' id='AddSkill' value='Add >' style='width: 80px;'><br>
            <input type='button' name='RemoveSkill' id='RemoveSkill' value='< Remove' style='width: 80px;'>
        </td>
        <td width='300' align='center'>
            <select name='cmbSkillsB' id='cmbSkillsB' size='10' style='width: 262px;'>
<%
                For b = 1 To j
%>
                    <option value='<%=TableB_ID(b)%>'><%=TableB_Description(b)%>
<%
                Next
%>
            </select>
        </td>
<!-- END -->
    </tr>
    <tr>
        <td colspan='3'>&nbsp;</td>
    </tr>
    <tr>
        <td colspan='3'><hr></td>
    </tr>
    <tr>
        <td colspan='3' align='center'><input type='button' name='Save' id='Save' Value='Save' style='width: 150px;'></td>
    </tr>
</table>
</form>
<%
Response.Write "<form action='users.asp' name='frmReturn'>" & vbcrlf
Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
Response.Write "<input type='hidden' name='Username' id='Username' value='" & Username & "'>" & vbcrlf
Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
Response.Write "</form>" & vbcrlf
Response.Write "<script language='vbscript'>" & vbcrlf
Response.Write "Sub xWindow_OnLoad" & vbcrlf
Response.Write "    msgbox ""Working on it...""" & vbcrlf
Response.Write "    frmReturn.Submit" & vbcrlf
Response.Write "End Sub" & vbcrlf
Response.Write "</script>" & vbcrlf
%>
</body>
</html>
