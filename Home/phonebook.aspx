<%@ Page Debug="true" %>
<!--#include virtual="includes/dataaccess.inc"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<%
Dim i, j
Dim User()
Dim Username
Dim StartTime, EndTime, HoursPerDay, MinorInterval, IntervalsPerHour, NoOfPeriods
Dim ColWidth
Dim LoopTime, DoHour, CurrentPeriod
Dim Phone, PhoneNo
%>

<html>
<head>
	<link rel='stylesheet' type='text/css' href='style.css'>
    <title>Phonebook</title>
</head>

<body>
    <div align='center' class='Body'>
        Empired Phonebook
    </div>
    <br>
    <div align='left' class='Body'>
<%
  SQL = "Exec sp_Users"
  adoRS.Open (SQL,adoConn)%>

    <table border='0' cellborder='0' cellpadding='0' cellspacing='0' align="center">
      <tr>
        <td class="TableCol1">&nbsp;Name&nbsp;</td>
        <td class="TableCol1">&nbsp;Phone&nbsp;</td>
        <td class="TableCol1">&nbsp;Email&nbsp;</td>
        <td class="TableCol1">&nbsp;Edit&nbsp;</td>
      </tr>
<%  Do Until adoRS.eof%>
      <tr>
        <td width="150">&nbsp;<%=adoRS("CName").Value & " " & adoRS("SName").Value%>&nbsp;</td>
<%      ' Turn each phone number into the following format:
        ' 04xx xxx xxx
        Phone = Trim(adoRS("Mobile").Value)
        PhoneNo = Mid(Phone, 1, 4) & " " & Mid(Phone, 5 , 3) & " " & Mid(Phone, 8, 3)%>
        <td>&nbsp;<%=PhoneNo%>&nbsp;</td>
        <td align="center">&nbsp;<a href="mailto:<%=adoRS("CName").Value & "." & adoRS("SName").Value%>@empired.com"><img src="images/email.gif" alt="Email <%=adoRS("CName").Value%>" border="0"></a>&nbsp;</td>
        <td align="center">&nbsp;<a href="edit_user_details.aspx?Person=<%=adoRS("ID").Value%>"><img src="images/edit.gif" alt="Edit details" border="0"></a>&nbsp;</td>
      </tr>
<%  adoRS.MoveNext
  Loop
%>
    </table>
    </div>

</body>
</html>
