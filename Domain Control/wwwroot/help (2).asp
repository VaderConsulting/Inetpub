<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<%
    Dim Reference
    
    Reference = Request.QueryString("reference")
%>
<html>
<!--#include file="styles.inc"-->
<head>
	<title>Help</title>
</head>

<body>
<div align='center'>
<%
    Select Case Reference
        Case "start"
            Response.Write "<br>" & vbcrlf
            Response.Write "Select an option from the list above." & vbcrlf
            Response.Write "<br>" & vbcrlf
        Case Else
    End Select
%>
</div>


</body>
</html>
