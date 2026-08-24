<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<%
  Dim UserTo, UserFrom, Subject, Body
  Dim objSendMail
  
  UserTo = Request.QueryString("UserTo")
  UserFrom = Request.QueryString("UserFrom")
  Subject = Request.QueryString("Subject")
  Body = Request.QueryString("Body")
%>

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
  <title>SendEmail</title>
</head>

<body>
<%
    Set objSendMail =      CreateObject("CDO.Message")
    objSendMail.From =     UserFrom
    objSendMail.To =       UserTo
    objSendMail.Subject =  Subject
    objSendMail.TextBody = Body
    objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "mammoth01" 
    objSendMail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
    objSendMail.Configuration.Fields.Update
    objSendMail.Send
    Set objSendMail = Nothing
%>

Your message has been sent.

</body>
</html>
