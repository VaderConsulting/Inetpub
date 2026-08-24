<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
	<title>Process SQL Statement</title>
</head>
<%
    Dim ClientName, Operation, ClientCode, ClientID, ProjectID, NewTaskName, SelectedTaskName
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
    SelectedTaskName = Request("SelectedTaskName")
    TaskID  = Request("TaskID")
    Skillname = Request("Skillname")
    DivisionName = Request("DivisionName")
    Operation = Request("Operation")
    Tablename = Request("Tablename")
    Username = Request("Username")
    
    adoConn.Open Session("DSN")
%>
<body>
<%
    ' What Operation is it ?
    
    Select Case Operation
        Case "INSERT"
            ' What Table is it ?
            Select Case Tablename
                Case "tblClient"
                    '1. Does the Client name already exist ?
                    SQL = "SELECT Clientname, ClientCode FROM tblClient WHERE ClientName LIKE '" & ClientName & "'"
                    adoRS.Open SQL, adoConn
                    
                    If adoRS.EOF Then
                        ' No.  Add the client
                        SQL = "INSERT INTO tblClient (ClientName) VALUES ('" & ClientName & "')"
                        adoConn.Execute SQL
                        
                        SQL = "SELECT TOP 1 ClientName, ClientID FROM tblClient WHERE ClientName = '" & ClientName & "' ORDER BY ClientID DESC"
                        adoRS.Close
                        adoRS.Open SQL, adoConn
                        ClientID = adoRS("ClientID")
                        ClientCode = Left(adoRS("ClientName"),4) & ClientID
                        Response.Write "<script language='vbscript'>" & vbcrlf
                        Response.Write vbcrlf
                        Response.Write "    Sub Confirmed" & vbcrlf
                        Response.Write "        frmInput.SQL.Value=" & Chr(34) & "UPDATE tblClient SET ClientCode = '" & Chr(34) & " & Ucase(frmInput.ClientCode.Value) & " & Chr(34) & ClientID & Chr(34) & " & " & Chr(34) & "' WHERE ClientID = " & ClientID & Chr(34) & vbcrlf
                        Response.Write "        frmInput.Submit" & vbcrlf
                        Response.Write "    End Sub" & vbcrlf
                        Response.Write vbcrlf
                        Response.Write "</script>" & vbcrlf
                        Response.Write "<br>" & vbcrlf
                        Response.Write "The following is the suggested ClientCode for " & ClientName & ": " & ClientCode & ".<br>" & vbcrlf
                        Response.Write "You may modify the string portion of this ClientCode if you wish." & vbcrlf
                        Response.Write "<form action='confirmed.asp' method='post' name='frmInput' id='frmInput'>" & vbcrlf
                        Response.Write "    <input type='text' name='ClientCode' id='ClientCode' value='" & Left(adoRS("ClientName"),4) & "'><br>" & vbcrlf
                        Response.Write "    Modify as necessary, and select <input type='button' value='Add' onclick='Confirmed'> when ready." & vbcrlf
                        Response.Write "    <input type='hidden' name='SQL' id='SQL' value=''>" & vbcrlf
                        Response.Write "    <input type='hidden' name='Tablename' id='Tablename' value='" & Tablename & "'>" & vbcrlf
                        Response.Write "</form>" & vbcrlf
                        adoRS.Close
                        adoRS.Open SQL, adoConn
                        
                        'Response.Write "Client ID: " & adoRS("ClientID") & "<br>" & vbcrlf
                    Else
                        ' This Client name already exists.
                        Response.Write ClientName & " already exists. (ClientCode=" & adoRS("ClientCode") & ")" & vbcrlf
                        Response.Write "<br>" & vbcrlf
                        adoRS.Close
                    End If
                Case "tblProject"
                    '1. Does the Project name already exist ?
                    SQL = "SELECT ProjectName FROM tblProject WHERE ProjectName LIKE '" & ProjectName & "' AND ClientID = " & ClientID
                    adoRS.Open SQL, adoConn
                    
                    If adoRS.EOF Then
                        ' No.  Add the Project
                        SQL = "INSERT INTO tblProject (ProjectName, ClientID) VALUES ('" & ProjectName & "'," & ClientID & ")"
                        adoConn.Execute SQL
                        SQL = "SELECT ClientCode FROM tblClient WHERE ClientID = " & ClientID
                        'Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                        adoRS2.Open SQL, adoConn
                        ClientCode = adoRS2("ClientCode")
                        'Response.Write "<!--" & vbcrlf & "Client Code: " & ClientCode & vbcrlf & "-->" & vbcrlf
                        adoRS2.Close
                        SQL = "SELECT ProjectID FROM tblProject WHERE ProjectName LIKE '" & ProjectName & "' AND ClientID = " & ClientID
                        'Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                        adoRS2.Open SQL, adoConn
                        ProjectID = adoRS2("ProjectID")
                        adoRS2.Close
                        SQL = "UPDATE tblProject SET ProjectCode = '" & ClientCode & "-" & ProjectID & "' WHERE ProjectID = " & ProjectID
                        adoConn.Execute SQL
                        Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                        Response.Write "<div align='center'>" & vbcrlf
                        Response.Write "    " & Tablename & " table has been updated." & vbcrlf
                        Response.Write "</div>" & vbcrlf
                        Response.Write "<form action='projects.asp' name='frmReturn'>" & vbcrlf
                        Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='ProjectName' id='ProjectName' value='" & ProjectName & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                        Response.Write "</form>" & vbcrlf
                        Response.Write "<script language='vbscript'>" & vbcrlf
                        Response.Write "Sub Window_OnLoad" & vbcrlf
                        'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                        Response.Write "    frmReturn.Submit" & vbcrlf
                        Response.Write "End Sub" & vbcrlf
                        Response.Write "</script>" & vbcrlf
                    Else
                        ' This Project name already exists.
                        Response.Write ProjectName & " already exists." & vbcrlf
                        Response.Write "<br>" & vbcrlf
                        adoRS.Close
                    End If
                Case "tblDivision"
                    '1. Does the Division name already exist ?
                    SQL = "SELECT Name FROM tblDivision WHERE Name LIKE '" & DivisionName & "'"
                    adoRS.Open SQL, adoConn
                    
                    If adoRS.EOF Then
                        ' No.  Add the Division
                        SQL = "INSERT INTO tblDivision (Name) VALUES ('" & DivisionName & "')"
                        adoConn.Execute SQL
                        Response.Write "<div align='center'>" & vbcrlf
                        Response.Write "    " & Tablename & " table has been updated." & vbcrlf
                        Response.Write "</div>" & vbcrlf
                        Response.Write "<form action='divisions.asp' name='frmReturn'>" & vbcrlf
                        Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                        Response.Write "</form>" & vbcrlf
                        Response.Write "<script language='vbscript'>" & vbcrlf
                        Response.Write "Sub Window_OnLoad" & vbcrlf
                        'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                        Response.Write "    frmReturn.Submit" & vbcrlf
                        Response.Write "End Sub" & vbcrlf
                        Response.Write "</script>" & vbcrlf
                    Else
                        ' This Division name already exists.
                        Response.Write DivisionName & " already exists." & vbcrlf
                        Response.Write "<br>" & vbcrlf
                        adoRS.Close
                    End If
                Case "tblSkills"
                    '1. Does the Skill name already exist ?
                    SQL = "SELECT Description FROM tblSkills WHERE Description LIKE '" & SkillName & "'"
                    adoRS.Open SQL, adoConn
                    
                    If adoRS.EOF Then
                        ' No.  Add the Skill
                        SQL = "INSERT INTO tblSkills (Description) VALUES ('" & SkillName & "')"
                        adoConn.Execute SQL
                        Response.Write "<div align='center'>" & vbcrlf
                        Response.Write "    " & Tablename & " table has been updated." & vbcrlf
                        Response.Write "</div>" & vbcrlf
                        Response.Write "<form action='skills.asp' name='frmReturn'>" & vbcrlf
                        Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                        Response.Write "</form>" & vbcrlf
                        Response.Write "<script language='vbscript'>" & vbcrlf
                        Response.Write "Sub Window_OnLoad" & vbcrlf
                        'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                        Response.Write "    frmReturn.Submit" & vbcrlf
                        Response.Write "End Sub" & vbcrlf
                        Response.Write "</script>" & vbcrlf
                    Else
                        ' This Skill name already exists.
                        Response.Write SkillName & " already exists." & vbcrlf
                        Response.Write "<br>" & vbcrlf
                        adoRS.Close
                    End If
                Case "tblTask"
                    '1. Does the Task name already exist ?
                    SQL = "SELECT Description FROM tblTask WHERE Description LIKE '" & NewTaskName & "' AND ProjectID = " & ProjectID
                    adoRS.Open SQL, adoConn
                    
                    If adoRS.EOF Then
                        ' No.  Add the Task
                        SQL = "INSERT INTO tblTask (Description, ProjectID) VALUES ('" & NewTaskName & "'," & ProjectID & ")"
                        'Response.Write SQL & "<br>"
                        adoConn.Execute SQL
                        Response.Write "<div align='center'>" & vbcrlf
                        Response.Write "    " & Tablename & " table has been updated." & vbcrlf
                        Response.Write "</div>" & vbcrlf
                        Response.Write "<form action='edittask.asp' name='frmReturn'>" & vbcrlf
                        Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                        Response.Write "</form>" & vbcrlf
                        Response.Write "<script language='vbscript'>" & vbcrlf
                        Response.Write "Sub Window_OnLoad" & vbcrlf
                        'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                        Response.Write "    frmReturn.Submit" & vbcrlf
                        Response.Write "End Sub" & vbcrlf
                        Response.Write "</script>" & vbcrlf
                    Else
                        ' This Taskname already exists.
                        Response.Write NewTaskName & " already exists." & vbcrlf
                        Response.Write "<br>" & vbcrlf
                        adoRS.Close
                    End If
                Case "tblUsers"
                    '1. Does the User name already exist ?
                    SQL = "SELECT Username FROM tblUsers WHERE Username LIKE '" & Username & "'"
                    adoRS.Open SQL, adoConn
                    
                    If adoRS.EOF Then
                        ' No.  Add the user
                        SQL = "INSERT INTO tblUsers (Username) VALUES ('" & Username & "')"
                        adoConn.Execute SQL
                        Response.Write "<div align='center'>" & vbcrlf
                        Response.Write "    " & Tablename & " table has been updated." & vbcrlf
                        Response.Write "</div>" & vbcrlf
                        Response.Write "<form action='users.asp' name='frmReturn'>" & vbcrlf
                        Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                        Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                        Response.Write "</form>" & vbcrlf
                        Response.Write "<script language='vbscript'>" & vbcrlf
                        Response.Write "Sub Window_OnLoad" & vbcrlf
                        'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                        Response.Write "    frmReturn.Submit" & vbcrlf
                        Response.Write "End Sub" & vbcrlf
                        Response.Write "</script>" & vbcrlf
                        'SQL = "SELECT TOP 1 ClientName, ClientID FROM tblClient WHERE ClientName = '" & ClientName & "' ORDER BY ClientID DESC"
                        'adoRS.Close
                        'adoRS.Open SQL, adoConn
                        'ClientID = adoRS("ClientID")
                        'ClientCode = Left(adoRS("ClientName"),4) & ClientID
                        'Response.Write "<script language='vbscript'>" & vbcrlf
                        'Response.Write vbcrlf
                        'Response.Write "    Sub Confirmed" & vbcrlf
                        'Response.Write "        frmInput.SQL.Value=" & Chr(34) & "UPDATE tblClient SET ClientCode = '" & Chr(34) & " & Ucase(frmInput.ClientCode.Value) & " & Chr(34) & ClientID & Chr(34) & " & " & Chr(34) & "' WHERE ClientID = " & ClientID & Chr(34) & vbcrlf
                        'Response.Write "        frmInput.Submit" & vbcrlf
                        'Response.Write "    End Sub" & vbcrlf
                        'Response.Write vbcrlf
                        'Response.Write "</script>" & vbcrlf
                        'Response.Write "<br>" & vbcrlf
                        'Response.Write "The following is the suggested ClientCode for " & ClientName & ": " & ClientCode & ".<br>" & vbcrlf
                        'Response.Write "You may modify the string portion of this ClientCode if you wish." & vbcrlf
                        'Response.Write "<form action='confirmed.asp' method='post' name='frmInput' id='frmInput'>" & vbcrlf
                        'Response.Write "    <input type='text' name='ClientCode' id='ClientCode' value='" & Left(adoRS("ClientName"),4) & "'><br>" & vbcrlf
                        'Response.Write "    Modify as necessary, and select <input type='button' value='Add' onclick='Confirmed'> when ready." & vbcrlf
                        'Response.Write "    <input type='hidden' name='SQL' id='SQL' value=''>" & vbcrlf
                        'Response.Write "    <input type='hidden' name='Tablename' id='Tablename' value='" & Tablename & "'>" & vbcrlf
                        'Response.Write "</form>" & vbcrlf
                        'adoRS.Close
                        'adoRS.Open SQL, adoConn
                    Else
                        ' This Username already exists.
                        Response.Write UserName & " already exists." & vbcrlf
                        Response.Write "<br>" & vbcrlf
                        adoRS.Close
                    End If
                Case Else
                    Response.Write "<br>" & vbcrlf
                    Response.Write "<div align='center'>" & vbcrlf
                    Response.Write Operation & " on " & Tablename & " is not supported YET" & vbcrlf
                    Response.Write "</div>" & vbcrlf
            End Select
        Case "DELETE"
            ' What Table is it ?
            Select Case Tablename
                Case "tblClient"
                    SQL = "DELETE FROM " & Tablename & " WHERE Clientname = '" & ClientName & "'"
                    Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                    adoConn.Execute SQL
                    Response.Write "<br>" & vbcrlf
                    Response.Write "<div align='center'>" & vbcrlf
                    Response.Write Operation & " on " & Tablename & " successful" & vbcrlf
                    Response.Write "</div>" & vbcrlf
                    Response.Write "<form action='clients.asp' name='frmReturn'>" & vbcrlf
                    Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                    Response.Write "</form>" & vbcrlf
                    Response.Write "<script language='vbscript'>" & vbcrlf
                    Response.Write "Sub Window_OnLoad" & vbcrlf
                    'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                    Response.Write "    frmReturn.Submit" & vbcrlf
                    Response.Write "End Sub" & vbcrlf
                    Response.Write "</script>" & vbcrlf
                Case "tblProject"
                    SQL = "DELETE FROM " & Tablename & " WHERE Projectname = '" & ProjectName & "'"
                    Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                    adoConn.Execute SQL
                    Response.Write "<br>" & vbcrlf
                    Response.Write "<div align='center'>" & vbcrlf
                    Response.Write Operation & " on " & Tablename & " successful" & vbcrlf
                    Response.Write "</div>" & vbcrlf
                    Response.Write "<form action='projects.asp' name='frmReturn'>" & vbcrlf
                    Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                    Response.Write "</form>" & vbcrlf
                    Response.Write "<script language='vbscript'>" & vbcrlf
                    Response.Write "Sub Window_OnLoad" & vbcrlf
                    'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                    Response.Write "    frmReturn.Submit" & vbcrlf
                    Response.Write "End Sub" & vbcrlf
                    Response.Write "</script>" & vbcrlf
                Case "tblSkills"
                    SQL = "DELETE FROM " & Tablename & " WHERE Description = '" & SkillName & "'"
                    Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                    adoConn.Execute SQL
                    Response.Write "<br>" & vbcrlf
                    Response.Write "<div align='center'>" & vbcrlf
                    Response.Write Operation & " on " & Tablename & " successful" & vbcrlf
                    Response.Write "</div>" & vbcrlf
                    Response.Write "<form action='skills.asp' name='frmReturn'>" & vbcrlf
                    Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                    Response.Write "</form>" & vbcrlf
                    Response.Write "<script language='vbscript'>" & vbcrlf
                    Response.Write "Sub Window_OnLoad" & vbcrlf
                    'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                    Response.Write "    frmReturn.Submit" & vbcrlf
                    Response.Write "End Sub" & vbcrlf
                    Response.Write "</script>" & vbcrlf
                Case "tblTask"
                    'SQL = "DELETE FROM " & Tablename & " WHERE Description = '" & SelectedTaskName & "' AND ProjectID = " & ProjectID
                    SQL = "DELETE FROM " & Tablename & " WHERE TaskID = " & TaskID
                    Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                    adoConn.Execute SQL
                    Response.Write "<br>" & vbcrlf
                    Response.Write "<div align='center'>" & vbcrlf
                    Response.Write Operation & " on " & Tablename & " successful" & vbcrlf
                    Response.Write "</div>" & vbcrlf
                    Response.Write "<form action='tasks.asp' name='frmReturn'>" & vbcrlf
                    Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='TaskID' id='TaskID' value='" & TaskID & "'>" & vbcrlf
                    Response.Write "</form>" & vbcrlf
                    Response.Write "<script language='vbscript'>" & vbcrlf
                    Response.Write "Sub Window_OnLoad" & vbcrlf
                    'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                    Response.Write "    frmReturn.Submit" & vbcrlf
                    Response.Write "End Sub" & vbcrlf
                    Response.Write "</script>" & vbcrlf
                Case "tblDivision"
                    SQL = "DELETE FROM " & Tablename & " WHERE Name = '" & DivisionName & "'"
                    Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                    adoConn.Execute SQL
                    Response.Write "<br>" & vbcrlf
                    Response.Write "<div align='center'>" & vbcrlf
                    Response.Write Operation & " on " & Tablename & " successful" & vbcrlf
                    Response.Write "</div>" & vbcrlf
                    Response.Write "<form action='divisions.asp' name='frmReturn'>" & vbcrlf
                    Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                    Response.Write "</form>" & vbcrlf
                    Response.Write "<script language='vbscript'>" & vbcrlf
                    Response.Write "Sub Window_OnLoad" & vbcrlf
                    'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                    Response.Write "    frmReturn.Submit" & vbcrlf
                    Response.Write "End Sub" & vbcrlf
                    Response.Write "</script>" & vbcrlf
                Case "tblUsers"
                    SQL = "DELETE FROM " & Tablename & " WHERE Username = '" & Username & "'"
                    Response.Write "<!--"  & vbcrlf & "SQL Statement: " & SQL & vbcrlf & "-->" & vbcrlf
                    adoConn.Execute SQL
                    Response.Write "<br>" & vbcrlf
                    Response.Write "<div align='center'>" & vbcrlf
                    Response.Write Operation & " on " & Tablename & " successful" & vbcrlf
                    Response.Write "</div>" & vbcrlf
                    Response.Write "<form action='users.asp' name='frmReturn'>" & vbcrlf
                    Response.Write "<input type='hidden' name='clientID' id='clientID' value='" & ClientID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='Username' id='Username' value='" & Username & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='ProjectID' id='ProjectID' value='" & ProjectID & "'>" & vbcrlf
                    Response.Write "<input type='hidden' name='SelectedTaskname' id='SelectedTaskname' value='" & SelectedTaskname & "'>" & vbcrlf
                    Response.Write "</form>" & vbcrlf
                    Response.Write "<script language='vbscript'>" & vbcrlf
                    Response.Write "Sub Window_OnLoad" & vbcrlf
                    'Response.Write "    msgbox ""Click <OK> to return to Tasks.""" & vbcrlf
                    Response.Write "    frmReturn.Submit" & vbcrlf
                    Response.Write "End Sub" & vbcrlf
                    Response.Write "</script>" & vbcrlf
                Case Else
                    Response.Write "<br>" & vbcrlf
                    Response.Write "<div align='center'>" & vbcrlf
                    Response.Write Operation & " on " & Tablename & " is not supported YET" & vbcrlf
                    Response.Write "</div>" & vbcrlf
            End Select
        Case Else
            Response.Write "<br>" & vbcrlf
            Response.Write "<div align='center'>" & vbcrlf
            Response.Write Operation & " on " & Tablename & " is not supported YET" & vbcrlf
            Response.Write "</div>" & vbcrlf
    End Select
    
    'Cleanup
    adoConn.Close
    Set adoRS=Nothing
    Set adoConn=Nothing
%>
</body>
</html>
