<% response.expires = -1 %>

<%
  SrvName = UCase(Request.QueryString("hostname"))
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<!--#include file="styles.inc"-->
<head>
	<title>Server Report for <%= SrvName %></title>
</head>
<body BGCOLOR="White">
  <DIV ALIGN="CENTER"><H1><%= SrvName %></H1>
      <TABLE WIDTH="100%">
<%            on error resume next
              MaxCol = 4
              ColCount = 1
              Dim Srv, ConnectString, S
              ConnectString = "WinNT://" + Srvname + ",COMPUTER"
              %><!-- ADSI Connectstring = <%=ConnectString%>-->
              <%
              Set Srv=GetObject(ConnectString)
              For Each Object in Srv%>
              <TR>
                <TD><%=Object.Name%></TD>
                <TD><%=Object.Class%></TD>
              </TR><%
              Next
              Set Srv = Nothing%>
            </TABLE>

        <TABLE WIDTH="100%">
<%        On Error Resume Next
          MaxCol = 4
          ColCount = 1
          ConnectString = "WinNT://" + Srvname + "/LANMANSERVER"
%>        <!-- ADSI Connectstring = <%=ConnectString%>-->
<%        Set Srv=GetObject(ConnectString)
          For Each Object in Srv%>
          <TR>
            <TD><%=Object.Name%></TD>
            <TD><%=Object.Path%></TD>
          </TR><%
          Next
          On Error Goto 0
          Set Srv = Nothing%>
        </TABLE>

      <TABLE WIDTH="100%">
<%      MaxCol = 4
        ColCount = 1
        Dim ServiceStatus(8)
        ServiceStatus(0) = "-UNKNOWN-"
        ServiceStatus(1) = "STOPPED"
        ServiceStatus(2) = "START_PENDING"
        ServiceStatus(3) = "STOP_PENDING"
        ServiceStatus(4) = "RUNNING"
        ServiceStatus(5) = "CONTINUE_PENDING"
        ServiceStatus(6) = "PAUSE_PENDING"
        ServiceStatus(7) = "PAUSED"
        ServiceStatus(8) = "ERROR"
        ConnectString = "WinNT://" + Srvname
%><!--  ADSI Connectstring = <%=ConnectString%>-->
<%      Set Srv=GetObject(ConnectString)
        Srv.Filter = Array("Service")
        For Each Object in Srv%>
        <TR>
          <TD><%=Object.DisplayName%></TD>
          <TD><%=ServiceStatus(Object.Status)%></TD>
        </TR><%
        Next
        Set Srv = Nothing%>
      </TABLE>
</DIV>
</body>
</html>
<%end if 

Function FieldExists(xField)
  On Error Resume Next
  i = CSCDB1(xField)
  FieldExists = (Err.Number = 0)
  Err.Clear
end Function
%>