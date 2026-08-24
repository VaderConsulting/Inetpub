<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<!--#include file="styles.inc"-->
<head>
	<title>Left Toolbar</title>
</head>

<body>

<%
    Dim ReferringPage
    ReferringPage = Request.Querystring("ReferringPage")
%>

<br>
<div align='center'>

<%
    Select Case ReferringPage
        Case "Computersearch"
            Response.Write "<div align='center'>" & vbcrlf
            Response.Write "<form name='frmSearch'>" & vbcrlf
            Response.Write "Search&nbsp;<input type='text' name='Hostname' id='Hostname' size='15' maxlength='15'><br>" & vbcrlf
            Response.Write "</form>" & vbcrlf
            Response.Write "</div>" & vbcrlf
            Response.Write "<table>" & vbcrlf
            Response.Write "    <tr>" & vbcrlf
            Response.Write "        <td width='20'><image src = 'images/windows_2000.gif' alt = 'Windows 2000' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/windows_nt.gif' alt = 'Windows NT' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/windows_xp.gif' alt = 'Windows XP' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/windows_unknown.gif' alt = 'Unknown Windows Version' width='18'></td>"
            Response.Write "    </tr>" & vbcrlf
            Response.Write "    <tr>" & vbcrlf
            Response.Write "        <td width='20'><image src = 'images/pdc.gif' alt = 'PDC' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/dc.gif' alt = 'BDC' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/netprinter.gif' alt = 'Print Server' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/server_database.gif' alt = 'SQL Server' width='18'></td>"
            Response.Write "    </tr>" & vbcrlf
            Response.Write "    <tr>" & vbcrlf
            Response.Write "        <td width='20'><image src = 'images/server_browser.gif' alt = 'Browse Server' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/server_ts.gif' alt = 'Terminal Server' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/cluster.gif' alt = 'Cluster Server' width='18'></td>"
            Response.Write "        <td width='20'><image src = 'images/sun.gif' alt = 'Unix' width='18'></td>"
            Response.Write "    </tr>" & vbcrlf
            Response.Write "    <tr>" & vbcrlf
            Response.Write "        <td align='right'><image src = 'images/time-server.gif' alt = 'Time Server' width='18'></td>"
            Response.Write "        <td align='right'><image src = 'images/workstation.gif' alt = 'Workstation' width='18'></td>"
            Response.Write "        <td align='right'><image src = 'images/server.gif' alt = 'Server' width='18'></td>"
            Response.Write "        <td align='right'><image src = 'images/advanced_server.gif' alt = 'Advanced Server' width='18'></td>"
            Response.Write "    </tr>" & vbcrlf
            Response.Write "</table>" & vbcrlf
        Case "Computers"
            Response.Write "<div align='center'>" & vbcrlf
            Response.Write "<form name='frmSearch'>" & vbcrlf
            Response.Write "Search&nbsp;<input type='text' name='Hostname' id='Hostname' size='15' maxlength='15'><br>" & vbcrlf
            Response.Write "</form>" & vbcrlf
            Response.Write "</div>" & vbcrlf
            Response.Write "<div align='left'>" & vbcrlf
            Response.Write "&nbsp;Add<br>" & vbcrlf
            Response.Write "</div>" & vbcrlf
        Case "ComputerInfo"
            Response.Write "<div align='left'>" & vbcrlf
            Response.Write "&nbsp;<img src='images/cpu.gif'>&nbsp;CPU Utilisation<br>" & vbcrlf
            Response.Write "&nbsp;<img src='images/delete.gif'>&nbsp;Delete<br>" & vbcrlf
            Response.Write "&nbsp;<img src='images/disk.gif'>&nbsp;Disk Space<br>" & vbcrlf
            Response.Write "&nbsp;<img src='images/ping.gif'>&nbsp;Ping<br>" & vbcrlf
            Response.Write "</div>" & vbcrlf
        Case "Users"
            Response.Write "Users<br>" & vbcrlf
            Response.Write "<div align='left'>" & vbcrlf
            Response.Write "<form name='frmSearch'>" & vbcrlf
            Response.Write "Search&nbsp;<input type='text' name='Username' id='Username' size='15' maxlength='15'><br>" & vbcrlf
            Response.Write "</form>" & vbcrlf
            Response.Write "</div>" & vbcrlf
        Case "Groups"
            Response.Write "Groups<br>" & vbcrlf
            Response.Write "<div align='left'>" & vbcrlf
            Response.Write "<form name='frmSearch'>" & vbcrlf
            Response.Write "Search&nbsp;<input type='text' name='Groupname' id='Groupname' size='15' maxlength='15'><br>" & vbcrlf
            Response.Write "</form>" & vbcrlf
            Response.Write "</div>" & vbcrlf
        Case "OUs"
            Response.Write "OU's<br>" & vbcrlf
            Response.Write "<div align='left'>" & vbcrlf
            Response.Write "<form name='frmSearch'>" & vbcrlf
            Response.Write "Search&nbsp;<input type='text' name='OUname' id='OUname' size='15' maxlength='15'><br>" & vbcrlf
            Response.Write "</form>" & vbcrlf
            Response.Write "</div>" & vbcrlf
        Case Else
            'Response.Write "Context sensitive<br>" & vbcrlf
            'Response.Write "Option 1<br>" & vbcrlf
            'Response.Write "Option 2<br>" & vbcrlf
            'Response.Write "Option 3<br>" & vbcrlf
            'Response.Write "Option 4<br>" & vbcrlf
            'Response.Write "Option 5<br>" & vbcrlf
            'Response.Write "Option 6<br>" & vbcrlf
            'Response.Write "Option 7<br>" & vbcrlf
            'Response.Write "Option 8<br>" & vbcrlf
            'Response.Write "Option 9<br>" & vbcrlf
            'Response.Write "Option 10<br>" & vbcrlf
    End Select
%>
</div>
<br>
<hr>
<br>
<div align='center'>
<img src='images/info.gif'><br>
<br>
<strong>Domain Status</strong><br>
<iframe src='status.asp' name='Status' id='Status' width='100%' height='150' marginwidth='0' marginheight='0' scrolling='no' frameborder='0' border="0">
</iframe>
</div>
</body>
</html>
