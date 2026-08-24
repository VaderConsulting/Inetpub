<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<!--#include file="styles.inc"-->
<head>
	<title>Toolbar</title>
</head>

<body>

<script language='vbscript'>

Sub ToolbarClick(SelectionNumber)
    Select Case SelectionNumber
    Case 1
        SelectionName = "<img width='19' src='images/warning.gif'>&nbsp;Alerts"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 2
        SelectionName = "<img width='19' src='images/workstation.gif'>&nbsp;Computers"
        parent.frames.item(1).window.location.href="left.asp?referringpage=Computersearch"
        parent.frames.item(2).window.location.href="main.asp?list=computersearch"
    Case 3
        SelectionName = "<img width='25' src='images/settings.ico'>&nbsp;Configuration"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 4
        SelectionName = "<img width='19' src='images/edit.gif'>&nbsp;Events"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 5
        SelectionName = "<img width='19' src='images/group.gif'>&nbsp;Groups"
        parent.frames.item(1).window.location.href="left.asp?referringpage=Groups"
        parent.frames.item(2).window.location.href="main.asp?list=groups"
    Case 6
        SelectionName = "<img width='19' src='images/card.gif'>&nbsp;Hardware"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 7
        SelectionName = "<img width='19' src='images/phone.gif'>&nbsp;Helpdesk"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 8
        SelectionName = "<img width='19' src='images/notes.gif'>&nbsp;Library"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 9
        SelectionName = "<img width='25' src='images/magnify.gif'>&nbsp;Monitor"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 10
        SelectionName = "<img width='19' src='images/ad.gif'>&nbsp;OU's"
        parent.frames.item(1).window.location.href="left.asp?referringpage=OUs"
        parent.frames.item(2).window.location.href="main.asp?list=ous"
    Case 11
        SelectionName = "<img width='19' src='images/report.gif'>&nbsp;Report"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 12
        SelectionName = "<img width='19' src='images/package.gif'>&nbsp;Software"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 13
        SelectionName = "<img width='19' src='images/info.gif'>&nbsp;Status"
        parent.frames.item(2).window.location.href="blank.htm"
    Case 14
        SelectionName = "<img width='19' src='images/user2.gif'>&nbsp;Users"
        parent.frames.item(1).window.location.href="left.asp?referringpage=Users"
        parent.frames.item(2).window.location.href="main.asp?list=users"
    Case 15
        SelectionName = ""
        parent.frames.item(2).window.location.href="blank.htm"
    End Select
    selection.innerhtml=SelectionName
End Sub

</script>

<table border='0'>
    <tr>
        <td width='130' rowspan=3 align='center'>&nbsp;Logo&nbsp;</td>
        <td width='270' rowspan=3 align='center'><strong><span id='selection'>Select a Menu Option</span></strong></td>
        <td width='100' align='center'><a onclick='ToolbarClick(1)' style='cursor : hand;'>Alerts</a></td>
        <td width='100' align='center'><a onclick='ToolbarClick(2)' style='cursor : hand;'>Computers</td>
        <td width='100' align='center'><a onclick='ToolbarClick(3)' style='cursor : hand;'>Configuration</td>
        <td width='100' align='center'><a onclick='ToolbarClick(4)' style='cursor : hand;'>Events</td>
        <td width='100' align='center'><a onclick='ToolbarClick(5)' style='cursor : hand;'>Groups</td>
    </tr>
    <tr>
        <td width='100' align='center'><a onclick='ToolbarClick(6)' style='cursor : hand;'>Hardware</td>
        <td width='100' align='center'><a onclick='ToolbarClick(7)' style='cursor : hand;'>Helpdesk</td>
        <td width='100' align='center'><a onclick='ToolbarClick(8)' style='cursor : hand;'>Library</td>
        <td width='100' align='center'><a onclick='ToolbarClick(9)' style='cursor : hand;'>Monitor</td>
        <td width='100' align='center'><a onclick='ToolbarClick(10)' style='cursor : hand;'>OU's</td>
    </tr>
        <td width='100' align='center'><a onclick='ToolbarClick(11)' style='cursor : hand;'>Report</td>
        <td width='100' align='center'><a onclick='ToolbarClick(12)' style='cursor : hand;'>Software</td>
        <td width='100' align='center'><a onclick='ToolbarClick(13)' style='cursor : hand;'>Status</td>
        <td width='100' align='center'><a onclick='ToolbarClick(14)' style='cursor : hand;'>Users</td>
        <td width='100' align='center'><a onclick='ToolbarClick(15)' style='cursor : hand;'>&nbsp;</td>
</table>

</body>
</html>
