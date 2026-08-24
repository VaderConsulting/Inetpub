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
