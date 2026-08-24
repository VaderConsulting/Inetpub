<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
	<title>Task Info</title>
</head>
<%
    'On Error Resume Next
    Dim ClientName, ClientCode, ProjectName, ProjectCode, StartDateTime, FinishDateTime, Internal, Chargable, ProjectComplete
    Dim TaskComplete, FixedPrice, MileStonePeriod, LastBillDateTime, WithinScope, Description, Rate, Billed, BillDateTime
    Dim NotSet
    Dim adoConn, adoRS, adoRS2
    Dim SQL
    Set adoConn = Server.CreateObject("ADODB.Connection")
    Set adoRS = Server.CreateObject("ADODB.Recordset")
    Set adoRS2 = Server.CreateObject("ADODB.Recordset")
    ClientName = Request("ClientName")
    ClientID = Request("ClientID")
    ProjectID = Request("ProjectID")
    ProjectName = Request("ProjectName")
    NewTaskName = Request("NewTaskName")
    TaskID = Request("TaskID")
    SelectedTaskName = Request("SelectedTaskName")
    'Skillname = Request("Skillname")
    'DivisionName = Request("DivisionName")
    'Operation = Request("Operation")
    'Tablename = Request("Tablename")
    'Username = Request("Username")
    
    adoConn.Open Session("DSN")
%>
<body>
<%
SQL =       "SELECT "
SQL = SQL & "tblClient.ClientName, tblClient.ClientCode, tblProject.ProjectName, tblProject.ProjectCode, tblProject.StartDateTime, tblProject.FixedPriceAmount, "
SQL = SQL & "tblProject.FinishDateTime, tblProject.Internal, tblProject.Chargable, tblProject.Complete AS ProjectComplete, tblProject.FixedPrice, tblProject.Billed AS ProjectBilled, "
SQL = SQL & "tblProject.MilestonePeriod, tblProject.LastBillDateTime, tblTask.WithinScope, tblTask.Description, tblTask.Rate, tblTask.Billed AS TaskBilled, "
SQL = SQL & "tblTask.BillDateTime, tblTask.Complete AS TaskComplete, tblTask.isMilestone "
SQL = SQL & "FROM (tblClient INNER JOIN tblProject ON tblClient.ClientID = tblProject.ClientID) INNER JOIN tblTask ON tblProject.ProjectID = tblTask.ProjectID "
SQL = SQL & "WHERE tblTask.TaskID = " & TaskID

Response.Write "<!--" & vbcrlf & SQL & vbcrlf & "-->" & vbcrlf
adoRS2.Open SQL, adoConn
If Not adoRS2.EOF Then
    ClientName = adoRS2("ClientName")
    ClientCode = adoRS2("ClientCode")
    ProjectName = adoRS2("ProjectName")
    ProjectCode = adoRS2("ProjectCode")
    StartDateTime = adoRS2("StartDateTime")
    FinishDateTime = adoRS2("FinishDateTime")
    Internal = adoRS2("Internal")
    Chargable = adoRS2("Chargable")
    ProjectComplete = adoRS2("ProjectComplete")
    TaskComplete = adoRS2("TaskComplete")
    FixedPrice = adoRS2("FixedPrice")
    ProjectBilled = adoRS2("ProjectBilled")
    FixedPriceAmount = adoRS2("FixedPriceAmount")
    MilestonePeriod = adoRS2("MilestonePeriod")
    LastBillDateTime = adoRS2("LastBillDateTime")
    WithinScope = adoRS2("WithinScope")
    Description = adoRS2("Description")
    Rate = adoRS2("Rate")
    TaskBilled = adoRS2("TaskBilled")
    BillDateTime = adoRS2("BillDateTime")
    isMilestone = adoRS2("isMilestone")
End If
NotSet = "<span class='highlight'>&lt;NOT SET&gt;</span>"
If Rate = 0 Then Rate = "??"
If (StartDateTime = "") Or (IsNull(StartDateTime)) Then StartDateTime = NotSet
If (FinishDateTime = "") Or (IsNull(FinishDateTime)) Then FinishDateTime = NotSet
adoRS2.Close
%>
<form action='userinfo.asp' name='frmEdit'>
&nbsp;<strong>Client</strong>:&nbsp;<%=ClientName & " (" & ClientCode & ")"%><br>
&nbsp;<strong>Project</strong>:&nbsp;<%=ProjectName & " (" & ProjectCode & ")"%><br>
&nbsp;<strong>Task</strong>:&nbsp;<input type='text' name='description' id='description' value='<%=Description%>' style='width:250;'><br>
<hr>
&nbsp;<strong>Project Information</strong><br>
&nbsp;The client for this<%If Internal Then%> internal<%Else%> external<%End If%> project <%If TaskBilled or ProjectBilled Then%>was <%Else%>is <%End If%><%If Chargable And Not FixedPrice Then%> charged per hour for this Task.<%ElseIf Chargable And FixedPrice Then%>charged at a fixed price for this Project ($<%=FixedPriceAmount%>).<%ElseIf Not Chargable Then%> not to be charged.<%End If%><br>
&nbsp;This project<%If ProjectComplete Then%> is complete and progressed from <%=StartDateTime%> to <%=FinishDateTime%><%Else%> is not complete.  <%End If%><br>
<%If Not ProjectComplete Then%>&nbsp;It <%If StartDateTime < now Then%> was<%Else%> is<%End If%> due to start on <%=StartDateTime%><%If FinishDateTime < Now Then%>, and to finish by <%=FinishDateTime%><%End If%>.<%End If%>
<hr>
&nbsp;<strong>Task Information</strong><br>
&nbsp;This task <%If TaskBilled or ProjectBilled Then%>was <%Else%>is <%End If%><%If Chargable And Not FixedPrice Then%> charged at $<%=Rate%> per hour.<%ElseIf Chargable And FixedPrice Then%>charged at a fixed price as part of the project.<%ElseIf Not Chargable Then%> free of charge.<%End If%><br>
&nbsp;This task<%If TaskComplete Then%> is complete.<%Else%> is not complete.  <%End If%><br>
&nbsp;This task <%If isMilestone Then%><%If TaskComplete Then%>was a <%Else%>has been defined as a <%End If%><strong><%Else%> is not a <%End If%>milestone.</strong>
<hr>

<table border='0' cellspacing='0' cellpadding='0' align='center'>
    <tr>
        <td align='center'>Skill List</td>
        <td>&nbsp;</td>
        <td align='center'>Required Skills</td>
    </tr>
    <tr>
<!-- START -->
<%
        ' Separate skills into two tables:
        ' All Skills available to choose from, (TableA), and those required for this task (TableB).
        SQL =       "SELECT tblSkills.Description, tblSkills.SkillID, tblTask.TaskID "
        SQL = SQL & "FROM (tblSkills LEFT JOIN tblSkillsRequired ON tblSkills.SkillID = tblSkillsRequired.SkillID) LEFT JOIN tblTask ON tblSkillsRequired.TaskID = tblTask.TaskID "
        SQL = SQL & "ORDER BY tblSkills.Description"
        Response.Write "<!!- " & vbcrlf & SQL & vbcrlf & "-->"
        adoRS.Open SQL, adoConn
        Do Until adoRS.EOF
            If adoRS("TaskID") & "" = TaskID Then
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
                For b = 1 To b
%>
                    <option value='<%=TableB_ID(b)%>'><%=TableB_Description(b)%>
<%
                Next
%>
            </select>
        </td>
<!-- END -->
    </tr>
</table>
<br>
<div align='center'>
<input type='button' name='Save' id='Save' Value='Save' style='width: 150px;'>
</div>
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
