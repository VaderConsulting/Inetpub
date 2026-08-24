<%@ Page Debug="true" %>
<!--#include virtual="includes/dataaccess.inc"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<%
Dim i, j, a, b
Dim fso, Working(7), Workdays, Username, Mobile, Message
Dim Person, Submitted, Phone, PhoneNo, CName, SName

' Get page variables
  Person = Request.QueryString("Person")
  If Person = "" Then 
    Person = Request.form("Person")
  End If
  
' Get form variables
  Username = Request.form("Username")
  Mobile = Replace(Request.form("Mobile")," ","")
  Workdays = Request.form("Output")
  Submitted = Request.form("Submitted")
%>
<html>
<head>
	<link rel='stylesheet' type='text/css' href='style.css'>
  <title>Edit Person Details</title>
</head>

<body>

<script language="VBScript">
  Sub DoOutput
    For i = 2 to 8
      a = Window.Document.Forms("Details").Elements.Item(i).checked
      if a then 
        b = 1
      else
        b = 0
      end if
      c = c & b
    Next
    Details.Output.Value = CStr(c)
    Details.Username.Value = UCase(Details.Username.Value)
    
    '' Set variable so this page knows the values were submitted.
    Details.Submitted.value="submitted"
    Details.submit
  End Sub
</script>
<%
  If Submitted <> "" Then
    ' Use person FORM variable instead
    'Person = Request.form("Person")
    
    SQL = "EXEC sp_UpdateUser " & Person & ",'" & Username & "','" & Mobile & "','" & Workdays & "'"
    Message = "Person Details Updated"
    ' Save output to database
    adoConn.Execute (SQL)
    'Response.Write SQL & "<br>"
  End If

  on error resume next
  SQL = "EXEC SP_GetUser " & Person
  adoRS.Open (SQL, adoConn)
  ID = Person
  CName = adoRS("CName").Value
  SName = adoRS("SName").Value
  Username = adoRS("Username").Value
  WorkDays = adoRS("WorkDays").Value
  
  ' Get each workday
  For i = 0 to 6
    Working(i) = mid(Workdays,i,1)
  Next
  
  ' Turn each phone number into the following format:
  ' 04xx xxx xxx
  Phone = Trim(adoRS("Mobile").Value)
  PhoneNo = Mid(Phone, 1, 4) & " " & Mid(Phone, 5 , 3) & " " & Mid(Phone, 8, 3)
  %>
  <br>
  <br>
  <br>
  <form name="Details" id="Details" target="_self" method="post">
  <table border='0' cellborder='0' cellpadding='0' cellspacing='0' align="center">
    <tr>
      <td colspan="2" align="center" class="TableCol1"><%=CName & " " & SName%></td>
    </tr>
    <tr>
      <td width="100">Username</td>
      <td width="100"><input type="text" id="Username" name="Username" value="<%=UCase(Username)%>" style="width: 100px;"></td>
    </tr>
    <tr>
      <td width="100">Mobile</td>
      <td width="100"><input type="text" id="Mobile" name="Mobile" value="<%=PhoneNo%>" style="width: 100px;"></td>
    </tr>
    <tr>
      <td colspan="2">&nbsp;</td>
    </tr>
    <tr>
      <td colspan="2" align="center">
        <table border='0' cellborder='0' cellpadding='0' cellspacing='0' align="center">
          <tr>
            <td colspan="7" align="center">Work Days</td>
          </tr>
          <tr>
            <td width="15" align="center">S</td>
            <td width="15" align="center">M</td>
            <td width="15" align="center">T</td>
            <td width="15" align="center">W</td>
            <td width="15" align="center">T</td>
            <td width="15" align="center">F</td>
            <td width="15" align="center">S</td>
          </tr>
          <tr>
<%          For i = 1 to 7%>
                <td width="15" align="center"><input type="Checkbox"<%If Working(i) = "1" Then%> checked<%End If%>></td>
<%          Next%>
          </tr>
          <tr>
            <td colspan="7" align="center">&nbsp;</td>
          </tr>
          <tr>
            <td colspan=7 align="center"><input type="button" value="Update" onclick="DoOutput"></td>
          </tr>
        </table>
      </td>
    </tr>
  </table>
<input type="hidden" id="Output" name="Output">
<input type="hidden" id="Person" name="Person" value="<%=Person%>">
<input type="hidden" id="Submitted" name="Submitted">
</form>
<span id="Message"></span>
<%adoConn.Close
  adoRS = Nothing
  adoConn = Nothing
%>
<hr>
<a href="/home">Home</a>
</body>
</html>
