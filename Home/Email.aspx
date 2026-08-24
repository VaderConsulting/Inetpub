<!--#include virtual="includes/dataaccess.inc"--> 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<%
  Dim UserTo
  Dim UserFrom
  Dim MailDomain1, MailDomain2, MailDomain3
  Dim Message, Footer
  Dim CName, SName
  Dim EmailAddressFrom, EmailAddressTo, Subject, MessageFrom, Company, PreformedSubject
  Dim objSendMail
  
  UserTo = Request.QueryString("UserTo")
  UserFrom = Request.QueryString("UserFrom")
  CName = Request.QueryString("CName")
  SName = Request.QueryString("SName")
  
  MailDomain1 = "@EMPIRED.COM"
  MailDomain2 = "@TUSK.COM.AU"
  MailDomain3 = "@BIGREDSKY.COM"

  ' Get form variables
  EmailAddressFrom = Request.Form("EmailFrom")
  EmailAddressTo = Request.Form("EmailTo")
  Message = Request.Form("Details")
  Subject = Request.Form("EmailSubject")
  MessageFrom = Request.Form("MessageFrom")
  Company = Request.Form("Company")
  PreformedSubject = Request.form("PreformedSubject")
  
  ' This page submits to itself.
  ' If this page has been submitted, process it and send email.
  If Message <> "" Then
    Footer = vbcrlf & "This email has been created automagically using Empired Locator (http://tuskprint/home)"
    objSendMail =      CreateObject("CDO.Message")
    objSendMail.From =     EmailAddressFrom
    objSendMail.To =       EmailAddressTo & MailDomain1 & "; " & EmailAddressTo & MailDomain2 & "; " & EmailAddressTo & MailDomain3
    objSendMail.Subject =  Subject
    objSendMail.TextBody = MessageFrom & " from " & Company & PreformedSubject & vbcrlf & Message & Footer
    objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mammoth01" 
    objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objSendMail.Configuration.Fields.Update
    on error resume next
    objSendMail.Send
    on error goto 0
    objSendMail = Nothing
  %>
  <html>
  <head>
    <link rel="stylesheet" type="text/css" href="style.css">
    <title>Email Sent</title>
  </head>
  
  <body>
  <strong>The message:</strong><br>
  <hr>  
  From:&nbsp;<%=EmailAddressFrom%><br>
  To:&nbsp;<%=EmailAddressTo & MailDomain1 & "; " & EmailAddressTo & MailDomain2 & "; " & EmailAddressTo & MailDomain3%><br>
  <br>
  Subject:<%=Subject%><br>
  <br>
  Message:<br><%=Replace(Message & Footer,vbcrlf,"<br>")%><br>
  <br>
  <hr>
  <strong>has been queued for sending.</strong><br>
  <strong>Empired Locator relies upon SMTP mail, and so mail may not be sent or received immediately.</strong><br>
  <br>
  Close this window to return.
  </body>
  
  </html>
<%Else%>

<% ' Turn the userID into a username
  ' Usage of Stored Procedure: sp_GetUserFromID ID
  ' Returns CName, SName, Username
  
  SQL = "Exec sp_GetUserFromID " & UserTo
  adoRS.Open (SQL,adoConn)
  Userto = adoRS("CName").Value & "." &adoRS("SName").Value
  adoRS.Close
%>

<html>
<head>
	<META HTTP-EQUIV="Page-Enter" CONTENT="filter:RevealTrans(Duration=1.000, Transition=20)">
  <link rel="stylesheet" type="text/css" href="style.css">
  <title>Email</title>
</head>

<body>

<script language="VBScript">

Function MessageTypeClick(Clicked)
  Select Case Clicked
    Case 1
      Msg.InnerHTML = "left a phone message."
    Case 2
      Msg.InnerHTML = "dropped by and left a message."
    Case 3
      Msg.InnerHTML = "left a package for you at reception."
  End Select
End Function

Function FromSelected
  '' //Message.Details.value = "Date: " & Date & chr(13) & "Time: " & Time & chr(13)
End Function

Function SendMail
    PersonName = Window.Document.Forms("Message").Elements.Item(3).value
    CompanyName = Window.Document.Forms("Message").Elements.Item(4).value
    
    '' //Window.Document.Forms("Message").Elements.Item(9).value = Window.Document.Forms("Message").Elements.Item(3).value & " " & msg.InnerHTML
    Window.Document.Forms("Message").Elements.Item(9).value = PersonName & " from " & CompanyName & " " & msg.InnerHTML
    Message.Submit
End Function

</script>

  <form name="Message" target="_self" method="post">
    <div class="Body">
      <table border="0" cellpadding="0" cellspacing="0" align="left" width="450">
        <tr>
          <td align="center" colspan="2">Message</td>
        </tr>
        <tr>
          <td colspan="2"><hr></td>
        </tr>
        <tr>
          <td width="30">&nbsp;From&nbsp;</td>
          <td>&nbsp;<%=UserFrom%>&nbsp;</td>
        </tr>
        <tr>
          <td width="30">&nbsp;To&nbsp;</td>
          <td>&nbsp;<%=UserTo & MailDomain1%>&nbsp;(<%=CName & " " & SName%>)</td>
        </tr>
        <tr>
          <td colspan="2">
            <table border="0" cellpadding="0" cellspacing="0">
              <tr>
                <td colspan="3"><hr></td>
              </tr>
              <tr height="30" valign="bottom">
                <td width="150"><input type="radio" Name="MessageType" value="1" Title="Phone Message" onclick="MessageTypeClick(1)">Phone message</td>
                <td width="150"><input type="radio" Name="MessageType" value="2" Title="Personal Message" onclick="MessageTypeClick(2)">Someone dropped by</td>
                <td width="150"><input type="radio" Name="MessageType" value="3" Title="Delivery Message" onclick="MessageTypeClick(3)">Package</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td><%=CName%>,</td>
        </tr>
        <tr>
          <td colspan="2"><input name="MessageFrom" type="text" value="<Enter name>" onclick="FromSelected">&nbsp;from&nbsp;<input name="Company" type="text" value="<Company name>">&nbsp;<span id="Msg"></span></td>
        </tr>
        <tr>
          <td colspan="2" height="40" valign="bottom">Details:</td>
        </tr>
        <tr>
          <td colspan="2"><textarea name="details" rows="5" cols="80"></textarea>
        </tr>
        <tr>
          <td colspan="2"><hr></td>
        </tr>
        <tr>
          <td colspan="2"><input type="button" value="Send" onclick="SendMail"></td>
        </tr>
        
      </table>
      <br>
  </div>
  <input type="hidden" name="EmailTo" value="<%=UserTo%>">
  <input type="hidden" name="EmailFrom" value="<%=UserFrom%>">
  <input type="hidden" name="EmailSubject" value="">
  <input type="hidden" name="PreformedSubject" id = "PreformedSubject">
  </form>

</body>

</html>
<%
  End If
%>