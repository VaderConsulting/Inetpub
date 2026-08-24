<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<!--#include file="./include/db.inc"-->
<!--#include file="./include/global.inc"-->
<%
    Dim Hostname
    Dim HostID
    Dim ReportType
    Dim Hosts
    Dim HostIP
    Dim HostState
    Dim HostWidth
    Dim RouteName
    Dim TimeUp
    
    Hostname   = UCase(Request.QueryString("Hostname"))
    HostID     = UCase(Request.QueryString("HostID"))
    ReportType = UCase(Request.QueryString("Type"))
    HostWidth  = 20 ' width used for columns and path lines
    
    ' ReportType:
    ' 1 = Day
    ' 2 = Week
    ' 3 = Month
    ' 4 = Year
    ' 9 = All
%>
<html>
<head>
	<title>Reports</title>
</head>

<body>
<div align="center">
    <h2>Outages</h2>
    <hr>
<%  adoConn.Open DSN
    ' Get Outages for each host
    if Hostname <> "TOMS" then
        SQL =       "SELECT tblTrace.Result, tblTrace.EventID, tblEvents.Event, tblHosts.Name, tblEvents.Time, tblEvents.ID, tblHosts.ID AS HostID, tblEvents.ID AS EventID "
        SQL = SQL & "FROM tblHosts RIGHT JOIN (tblTrace RIGHT JOIN tblEvents ON tblTrace.EventID = tblEvents.ID) ON tblHosts.ID = tblEvents.HostID "
        SQL = SQL & "WHERE tblEvents.Event = 'down' "
        
        SELECT CASE ReportType
            Case 1
                SQL = SQL & "AND DATEDIFF(day,tblEvents.[Time],GETDATE()) <= 1 "%>
                <h3>In the last 24 hours</h3>
<%          Case 2
                SQL = SQL & "AND DATEDIFF(day,tblEvents.[Time],GETDATE()) <= 7 "%>
                <h3>In the last 7 days</h3>
<%          Case 3
                SQL = SQL & "AND DATEDIFF(month,tblEvents.[Time],GETDATE()) <= 1 "%>
                <h3>In the last Month</h3>
<%          Case 4
                SQL = SQL & "AND DATEDIFF(year,tblEvents.[Time],GETDATE()) <= 1 "%>
                <h3>In the last Year</h3>
<%          Case 9 %>
                <h3>All Results</h3>
<%      End Select
        If trim(Hostname) <> "" then
            SQL = SQL & " AND tblHosts.Name ='" & Hostname & "' "
        End if
        SQL = SQL & "ORDER BY tblEvents.ID"
    Else ' Get Outages for TOMS
        SQL = "SELECT *, 'TOMS' AS Name FROM tblTOMS "
        SELECT CASE ReportType
            Case 1
                SQL = SQL & "WHERE DATEDIFF(day,Time,GETDATE()) <= 1 "%>
                <h3>In the last 24 hours</h3>
<%          Case 2
                SQL = SQL & "WHERE DATEDIFF(day,Time,GETDATE()) <= 7 "%>
                <h3>In the last 7 days</h3>
<%          Case 3
                SQL = SQL & "WHERE DATEDIFF(month,Time,GETDATE()) <= 1 "%>
                <h3>In the last Month</h3>
<%          Case 4
                SQL = SQL & "WHERE DATEDIFF(year,Time,GETDATE()) <= 1 "%>
                <h3>In the last Year</h3>
<%          Case 9 %>
                <h3>All Results</h3>
<%      End Select
        SQL = SQL & "AND Event = 'down' ORDER BY id"
    End If
    adoRS.open SQL, adoConn%>
</div>
<%  Response.write "<!-- SQL String used: " & SQL & "-->" & vbcrlf%>

<%          ' Display each of the hosts along the route as GIF images separated by black lines (GIF's again)
            if not adoRS.eof then%>
            <table border="0" cellspacing="0" cellpadding="0" align="left">
                <tr>
                    <td align="center" class="smallestblackfont"><strong>Date&nbsp;</strong></td>
                    <td align="center" class="smallestblackfont"><strong>Time Down&nbsp;</strong></td>
                    <td align="center" class="smallestblackfont"><strong>Time Up&nbsp;</strong></td>
                    <td align="center" class="smallestblackfont"><strong>Outage Time&nbsp;</strong></td>
                    <td align="center" class="smallestblackfont"><strong>Host&nbsp;</strong></td>
                    <td align="center" class="smallestblackfont" colspan="30">
<%                      if Hostname <> "TOMS" Then %>
                            <strong>Route</strong>
<%                      Else %>
                            &nbsp;
<%                      End If %>
                    </td>
                </tr>
                <tr>
                    <td colspan='30'><hr size='1'></td>
                </tr>
<%          end if

            do until adoRS.eof
                i = i +1
                If Hostname <> "TOMS" Then
                    HostTime = adoRS("Time")
                    Result = adoRS("Result")
                    HostName = adoRS("Name")
                    HostIP = "Router"
                    HostID = adoRS("HostID")
                    EventID = adoRS("EventID")
                    DateDown = formatdatetime(adoRS("Time"),2)
                    SQL2 = ""
                    SQL2 = SQL2 & "SELECT top 1 Time "
                    SQL2 = SQL2 & "FROM tblEvents WHERE Event = 'up' AND HostID = " &  HostID & " AND ID > " & EventID & " "
                    SQL2 = SQL2 & "ORDER BY tblEvents.ID"
                    Response.write "<!-- SQL2 String used: " & SQL2 & "-->" & vbcrlf
                    'response.end
                    adoRS2.open SQL2, adoConn
                    if not adoRS2.eof then
                        TimeUp = adoRS2("Time")
                        TotalTime = datediff("n",HostTime, TimeUp) & " minute(s)"
                        if totaltime = "0 minute(s)" then
                            TotalTime = datediff("s",HostTime, TimeUp) & " seconds"
                        end if
                    else
                        TimeUp = "Still down"
                        TotalTime = datediff("n",HostTime, now) & " minutes*"
                        if totaltime = "0 minutes*" then
                            TotalTime = datediff("s",HostTime, now) & " seconds*"
                        end if
                    end if
                    adoRS2.close
                Else
                    HostTime = adoRS("Time")
                    DateDown = formatdatetime(adoRS("Time"),2)
                    SQL2 = ""
                    SQL2 = "SELECT top 1 Time FROM tblTOMS WHERE ID > " & adoRS("ID")
                    Response.write "<!-- SQL2 String used: " & SQL2 & "-->" & vbcrlf
                    adoRS2.open SQL2, adoConn
                    if not adoRS2.eof then
                        TimeUp = adoRS2("Time")
                        TotalTime = datediff("n",HostTime, TimeUp) & " minute(s)"
                        if totaltime = "0 minute(s)" then
                            TotalTime = datediff("s",HostTime, TimeUp) & " seconds"
                        end if
                    else
                        TimeUp = "Still down"
                        TotalTime = datediff("n",HostTime, now) & " minutes*"
                        if totaltime = "0 minutes*" then
                            TotalTime = datediff("s",HostTime, now) & " seconds*"
                        end if
                    end if
                    adoRS2.close
                    if adoRS("Event") = "down" then
                        Result = "TOMS:0"
                    Else
                        Result = "TOMS:1"
                    End If
                    HostIP = "TOMS"
                End If%>
                <tr align="center" height = "20" valign="top">
                    <td nowrap valign='middle' class="smallestblackfont"><%=DateDown%>&nbsp;</td>
                    <td nowrap valign='middle' class="smallestblackfont"><%=formatdatetime(HostTime,3)%>&nbsp;</td>
                    <td nowrap valign='middle' class="smallestblackfont"><%if TimeUp <> "Still down" then response.write formatdatetime(TimeUp,3) else response.write TimeUp end if%>&nbsp;</td>
                    <td nowrap valign='middle' class="smallestblackfont"><%=TotalTime%>&nbsp;</td>
                    <td nowrap valign='middle' class="smallestblackfont"><%=adoRS("Name")%></td>
                    <td width = "<%=HostWidth %>">&nbsp;</td>
                    <td class="smallestblackfont">
                        <img src="./images/server_monitor.gif" alt="Monitoring Server"><br>
                        Monitor
                    </td>
<%                  Route =Result & ""
                    Hosts = Split(Route,",",-1,1)
                    For lp = LBound(Hosts) to UBound(Hosts)
                        HostIP =  left(Hosts(lp),len(Hosts(lp))-2)
                        SQL2 = "SELECT TOP 1 Name FROM tblRouteNames where IP = '" & HostIP & "' ORDER BY Name"
                        adoRS2.open SQL2, adoConn
                        if not adoRS2.eof then
                            RouteName = adoRS2("Name")
                        end if
                        adoRS2.close
                        HostState = right(Hosts(lp),1)%>
                        <td align="center"><img src="./images/path.gif" width = "<%=HostWidth%>" height="18"></td>
                        <td align="center"><img src="./images/path.gif" width = "<%=HostWidth%>" height="18"></td>
<%                      if lp <> ubound(Hosts) then ' ie we are not at the last record yet, so display a router picture
                            if HostState = "1" then%>
                                <td class="smallestblackfont">
                                    <img src="./images/router2.gif" alt ="<%=HostIP%>"><br>
                                    <%=RouteName%>
                                </td>
<%                          Else%>
                                <td class="smallestblackfont">
                                    <img src="./images/problem_router2.gif" alt ="<%=HostIP%>"><br>
                                    <%=RouteName%>
                                </td>
<%                          end if
                        Else                        ' Yes, this is the last record, ie the actual host itself, so display a server picture
                            If HostState = "1" then%>
                                <td class="smallestblackfont" align="center">
                                    <img src="./images/server.gif" alt ="<%=Hostname%>"><br>
                                    <%=Hostname%>
                                </td>
<%                          Else%>
                                <td class="smallestblackfont" align="center">
                                    <img src="./images/problem_server.gif" alt ="<%=Hostname%>"><br>
                                    <%=Hostname%>
                                </td>
<%                          End if
                        End if
                    Next
                    Route = "" %>
                    <td width = "<%=HostWidth%>">&nbsp;</td>
                </tr>
                <tr>
                    <td colspan='30'><hr size='1'></td>
                </tr>
<%              adoRS.Movenext
            loop
            if i = 0 then%>
                No outages for the given period.
<%          else %>
                
<%          end if %>
    </table>
    <!--<font class="smallestblackfont">The period for any single outage will be a <u>minimum</u> of 5 seconds due to the monitoring interval.</font>-->
<%  adoRS.Close
    adoConn.Close %>
</body>
</html>
