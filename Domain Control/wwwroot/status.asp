<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<META HTTP-EQUIV="REFRESH" CONTENT='5'>
<!--#include file="styles.inc"-->
<head>
	<title>Domain Status</title>
</head>

<body>

<br>
<div align='left'>
<%
    Dim i
    Randomize Timer
    Response.Write "<table>" & vbcrlf
    Response.Write "    <tr>" & vbcrlf
    i = int(rnd(1) * 100)
    Response.Write "        <td nowrap class='smallblackfont'>&nbsp;<strong>" & i & "</strong></td>"
    If i <> 1 Then
        Response.Write "        <td nowrap class='smallblackfont'>Events</td>" & vbcrlf
    Else
        Response.Write "        <td nowrap class='smallblackfont'>Event</td>" & vbcrlf
    End If
    Response.Write "        <td nowrap style='cursor : hand;'><img src='images/edit.gif'></td>" & vbcrlf
    Response.Write "    </tr>" & vbcrlf
    i = int(rnd(1) * 100)
    Response.Write "        <td nowrap class='smallblackfont'>&nbsp;<strong>" & i & "</strong></td>"
    If i <> 1 Then
        Response.Write "        <td nowrap class='smallblackfont'>Hosts Down</td>" & vbcrlf
    Else
        Response.Write "        <td nowrap class='smallblackfont'>Host Down</td>" & vbcrlf
    End If
    Response.Write "        <td nowrap style='cursor : hand;'><img src='images/stop.gif'></td>" & vbcrlf
    Response.Write "    </tr>" & vbcrlf
    i = int(rnd(1) * 100)
    Response.Write "        <td nowrap class='smallblackfont'>&nbsp;<strong>" & i & "</strong></td>"
    If i <> 1 Then
        Response.Write "        <td nowrap class='smallblackfont'>Failed Backups&nbsp;</td>" & vbcrlf
    Else
        Response.Write "        <td nowrap class='smallblackfont'>Failed Backup&nbsp;</td>" & vbcrlf
    End If
    Response.Write "        <td nowrap style='cursor : hand;'><img src='images/problem_backup.gif'></td>" & vbcrlf
    Response.Write "    </tr>" & vbcrlf
    Response.Write "</table>" & vbcrlf
%>
</div>
</body>
</html>
