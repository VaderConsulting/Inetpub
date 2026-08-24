<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!--#include file="./include/db.inc"-->
<!--#include file="./include/global.inc"-->
<%
    Dim TOMSState
    Dim HostState
    Dim SQL2
    'on error resume next
%>
<META HTTP-EQUIV="REFRESH" CONTENT="30">
<html>
<head>
	<title>TOMS Monitoring</title>
</head>

<body>

<div align="center">
    <h1>Real-time status</h1>
</div>
<hr>
<%  adoconn.commandtimeout=360
    adoConn.Open DSN
    
    'Report.asp:
    ' 1 = Day
    ' 2 = Week
    ' 3 = Month
    ' 4 = Year
    ' 9 = All%>

    <table>
        <tr>
            <td class="SmallBlackFont" align="center"><b>&nbsp;Hostname&nbsp;</b></td>
            <td class="SmallBlackFont" align="center"><b>&nbsp;Status&nbsp;</b></td>
            <td class="SmallBlackFont" align="center" colspan="2"><b>&nbsp;Outages<font class="smallredfont">*</font>&nbsp;</b></td>
            <td width ="30">&nbsp;</td>
            <td class="SmallBlackFont" align="center"><b>&nbsp;Reports&nbsp;</b></td>
        </tr>
        <tr>
<%  ' Get current state of TOMS Server
    SQL = "SELECT TOP 1 * FROM tblTOMS ORDER BY ID DESC"
    adoRS.open SQL, adoConn
    Do Until adoRS.EOF
        TOMSState = UCASE(adoRS("Event"))
        response.write "<!-- TOMS: " & TOMSState & "-->" & vbcrlf
        if TOMSState = "OK" then%>
            <td align="left" class="SmallestBlackFont">TOMS</td>
            <td align="center"><img src="./images/server_database.gif" alt="TOMS is Up"></td>
<%      else%>
            <td align="left" class="SmallestBlackFont"><strong>TOMS</strong></td>
            <td align="center"><img src="./images/problem_server_database.gif" alt="TOMS is Down"></td>
<%      end if
        ' Count the number of outages for TOMS in the last 24 hours
        SQL2 = "SELECT Count (*) AS x FROM tblTOMS WHERE Event = 'down' AND (DATEDIFF(hour, [Time], GETDATE()) <= 24)"
        response.write "<!-- SQL for Count1: " & SQL2 & "-->" & vbcrlf
        adoRS2.open SQL2, adoConn
        Count1 =0
        if not adoRS2.eof then 
            Count1 = adors2("x")
            response.write "<!-- Count1: " & Count1 & "-->" & vbcrlf
        end if
        adoRS2.Close
        
        ' Count the number of outages for TOMS in the last 1 hour
        SQL2 = "SELECT Count (*) AS x FROM tblTOMS WHERE Event = 'down' AND (DATEDIFF(hour, [Time], GETDATE()) <= 1)"
        response.write "<!-- SQL for Count2: " & SQL2 & "-->" & vbcrlf
        adoRS2.open SQL2, adoConn
        Count2 =0
        if not adoRS2.eof then 
            Count2 = adors2("x")
            response.write "<!-- Count2: " & Count2 & "-->" & vbcrlf
        end if
        adoRS2.Close%>
            <td align="center">&nbsp;<%=Count1%>&nbsp;</td>
            <td align="center">&nbsp;<%=Count2%>&nbsp;</td>
            <td width ="30">&nbsp;</td>
            <td align="center" nowrap>
                <a href="report.asp?type=1&hostname=TOMS" target="_blank"><img border="0" src="./images/sched_daily.gif" alt="TOMS outages in the last 24 hours"></a>
                <a href="report.asp?type=2&hostname=TOMS" target="_blank"><img border="0" src="./images/sched_weekly.gif" alt="TOMS outages in the last 7 days"></a>
                <a href="report.asp?type=3&hostname=TOMS" target="_blank"><img border="0" src="./images/sched_monthly.gif" alt="TOMS outages in the last month"></a>
                <a href="report.asp?type=4&hostname=TOMS" target="_blank"><img border="0" src="./images/sched_yearly.gif" alt="TOMS outages in the last year"></a>
                <a href="report.asp?type=9&hostname=TOMS" target="_blank"><img border="0" src="./images/schedule.gif" alt="All TOMS outages"></a>
            </td>
<%        adoRS.Movenext
    Loop
    adoRS.Close
%>
        </tr>
<%  ' Get state of each Host
    SQL = "SELECT * FROM tblHosts ORDER BY Name ASC"
    adoRS.open SQL, adoConn
    Do Until adoRS.EOF%>
        <tr>
<%      HostState = adoRS("isUp")
        if HostState = true then%>
            <td align="left" class="SmallestBlackFont"><%=adoRS("Name")%></td>
            <td align="center"><img src="./images/server.gif" alt ="<%=adoRS("Name")%> is up"></td>
<%      else%>
            <!-- <a href="route_info.asp?hostname=<%=adoRS("Name") & "&hostid=" &adoRS("ID")%>" target="_blank"> -->
            <td align="left"class="SmallestBlackFont"><strong><%=adoRS("Name")%></strong></td>
            <td align="center"><img src="./images/problem_server.gif" alt ="<%=adoRS("Name")%> is down""></td>
<%      end if
        ' Count the number of outages for the selected host in the last 24 hours
        SQL2 = "SELECT Count (*) AS x FROM tblEvents WHERE Event = 'down' AND HostID = " & adoRS("ID") & " AND (DATEDIFF(hour, [Time], GETDATE()) <= 24)"
        response.write "<!-- SQL for Count1: " & SQL2 & "-->" & vbcrlf
        adoRS2.open SQL2, adoConn
        Count1 =0
        if not adoRS2.eof then 
            Count1 = adors2("x")
            response.write "<!-- Count1: " & Count1 & "-->" & vbcrlf
        end if
        adoRS2.Close
        
        ' Count the number of outages for the selected host in the last 1 hour
        SQL2 = "SELECT Count (*) AS x FROM tblEvents WHERE Event = 'down' AND HostID = " & adoRS("ID") & " AND (DATEDIFF(hour, [Time], GETDATE()) <= 1)"
        response.write "<!-- SQL for Count2: " & SQL2 & "-->" & vbcrlf
        adoRS2.open SQL2, adoConn
        Count2 =0
        if not adoRS2.eof then 
            Count2 = adors2("x")
            response.write "<!-- Count2: " & Count2 & "-->" & vbcrlf
        end if
        adoRS2.Close%>
            <td align="center"><%=Count1%></td>
            <td align="center"><%=Count2%></td>
            <td width ="30">&nbsp;</td>
            <td align="center" nowrap>
                <a href="report.asp?type=1&hostname=<%=adoRS("Name")%>" target="_blank"><img border="0" src="./images/sched_daily.gif" alt="Outages for <%=adoRS("Name")%> in the last 24 hours"></a>
                <a href="report.asp?type=2&hostname=<%=adoRS("Name")%>" target="_blank"><img border="0" src="./images/sched_weekly.gif" alt="Outages for <%=adoRS("Name")%> in the last 7 days"></a>
                <a href="report.asp?type=3&hostname=<%=adoRS("Name")%>" target="_blank"><img border="0" src="./images/sched_monthly.gif" alt="Outages for <%=adoRS("Name")%> in the last month"></a>
                <a href="report.asp?type=4&hostname=<%=adoRS("Name")%>" target="_blank"><img border="0" src="./images/sched_yearly.gif" alt="Outages for <%=adoRS("Name")%> in the last year"></a>
                <a href="report.asp?type=9&hostname=<%=adoRS("Name")%>" target="_blank"><img border="0" src="./images/schedule.gif" alt="All outages for <%=adoRS("Name")%>"></a>
            </td>
        </tr>
<%      adoRS.Movenext
    Loop
    adoRS.Close%>
    </table>
<%  adoConn.Close%>
<font class ="smallestblackfont">(<font class="smallredfont">*</font>&nbsp;In the last 24 and 1 hour)</font><br>
<br>
<hr>
Select one of the following icons for outages over the given period, for all hosts.<br>
<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
&nbsp;<a href="report.asp?type=1" target="_blank"><img border="0" src="./images/sched_daily.gif" alt="Outages in the last 24 hours"></a>
&nbsp;<a href="report.asp?type=2" target="_blank"><img border="0" src="./images/sched_weekly.gif" alt="Outages in the last week"></a>
&nbsp;<a href="report.asp?type=3" target="_blank"><img border="0" src="./images/sched_monthly.gif" alt="Outages in the last month"></a>
&nbsp;<a href="report.asp?type=4" target="_blank"><img border="0" src="./images/sched_yearly.gif" alt="Outages in the last year"></a>
&nbsp;<a href="report.asp?type=9" target="_blank"><img border="0" src="./images/schedule.gif" alt="All outages"></a>
<!--<a href="add.asp" target="_blank">Add</a> or <a href="edit.asp" target="_blank">Edit</a> a Host or <a href="route.asp" target="_blank">Route</a>.<br>
<hr>-->
<!--#include file="./include/footer.inc"-->
</body>
</html>
