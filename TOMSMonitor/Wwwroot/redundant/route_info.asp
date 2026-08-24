<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!--#include file="./include/db.inc"-->
<!--#include file="./include/global.inc"-->
<%
    response.expires = -1
    Dim Hostname
    Dim HostID
    Dim Route
    Dim Hosts
    Dim HostIP
    Dim HostState
    Dim HostWidth
    
    Hostname = UCase(Request.QueryString("Hostname"))
    HostID   = UCase(Request.QueryString("HostID"))
    HostWidth = 35 ' width used for columns and path lines
%>
<html>
<head>
	<title>Routes</title>
</head>

<body>
<div align="center">
<h2>Route info for <%=Hostname%></h2>
</div>

<hr>
<%  adoConn.Open DSN
    on error resume next
    ' Get Trace Results for most recent outage
    SQL = "SELECT TOP 1 tblTrace.ID, tblTrace.Result, tblEvents.Time FROM tblTrace INNER JOIN "
    SQL = SQL & "(tblEvents INNER JOIN tblHosts ON tblEvents.HostID = tblHosts.ID) ON "
    SQL = SQL & "tblTrace.EventID = tblEvents.ID "
    SQL = SQL & "WHERE tblHosts.ID=" & HostID & " ORDER BY tblTrace.ID DESC"
    adoRS.open SQL, adoConn
    Response.write "<!-- SQL String used: " & SQL & "-->" & vbcrlf%>

<%          ' Display each of the hosts along the route as GIF images separated by black lines (GIF's again)%>

<%          If Not(adoRS.eof) Then%>
                <table border="0" cellspacing="0" cellpadding="0" align="center">
                <tr align="center">
                    <td width = "<%=HostWidth %>">&nbsp;</td>
                    <td><img src="./images/server_monitor.gif" alt="Monitoring Server"></td>
<%              Route =adoRS("Result") & ""
                Hosts = Split(Route,",",-1,1)
                For lp = LBound(Hosts) to UBound(Hosts)
                    HostIP =  left(Hosts(lp),len(Hosts(lp))-2)
                    HostState = right(Hosts(lp),1)%>
                    <td align="center"><img src="./images/path.gif" width = "<%=HostWidth%>" height="18"></td>
                    <td align="center"><img src="./images/path.gif" width = "<%=HostWidth%>" height="18"></td>
<%                  if lp <> ubound(Hosts) then ' ie we are not at the last record yet, so display a router picture
                        if HostState = "1" then%>
                            <td><img src="./images/router2.gif" alt ="<%=HostIP%>"></td>
<%                      Else%>
                            <td><img src="./images/problem_router2.gif" alt ="<%=HostIP%>"></td>
<%                      end if
                    Else                        ' Yes, this is the last record, ie the actual host itself, so display a server picture
                        if HostState = "1" then%>
                            <td><img src="./images/server.gif" alt ="<%=Hostname%>"></td>
<%                      Else%>
                            <td><img src="./images/problem_server.gif" alt ="<%=Hostname%>"></td>
<%                      end if
                    End if
                Next
                adoRS.Movenext
                Route = ""%>
            <td width = "<%=HostWidth%>">&nbsp;</td>
        </tr>
<%          Else%>
                <div align="center">
                    <b>There are no records for <%=Hostname%>.</b>
                </div>
<%          End If
            adors.movefirst%>
        <tr align="center" nowrap>
        
<%      ' Display just the route names (ie the IP Addresses)
        If Not(adoRS.eof) Then
            Route =adoRS("Result") & ""
            Hosts = Split(Route,",",-1,1)%>
            <td colspan="3" class="smallblackfont" width = "<%=HostWidth * 2%>" nowrap align="center">IP Address</td>
<%          For lp = LBound(Hosts) to UBound(Hosts)
                HostIP =  left(Hosts(lp),len(Hosts(lp))-2)%>
                <td colspan="3" class="smallblackfont"><%=HostIP%></td>
<%          Next
            adoRS.Movenext%>
            </tr>
<%          Route = ""
        End If%>

<%        adors.movefirst%>
        <tr align="center">
<%      ' Display just the route times
        If Not(adoRS.eof) Then
            Route =adoRS("Result") & ""
            Hosts = Split(Route,",",-1,1)%>
            <td colspan="3" class="smallestblackfont" width = "<%=HostWidth  * 2%>" nowrap align="center">Last Contact</td>
<%          For lp = LBound(Hosts) to UBound(Hosts)
                HostIP =  left(Hosts(lp),len(Hosts(lp))-2)
                HostState = right(Hosts(lp),1)
                HostTime = adoRS("Time")%>
                <td colspan="3" class="smallestblackfont"><%=day(HostTime) & "-" & month(HostTime) & "-" & year(HostTime) & " " & formatdatetime(HostTime,4)%></td>
<%          Next
            adoRS.Movenext%>
            </tr>
<%          Route = ""
        End If%>

    </table>
<%  adoRS.Close
    adoConn.Close%>
<!--#include file="./include/footer.inc"-->
</body>
</html>
