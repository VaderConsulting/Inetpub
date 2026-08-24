<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Menu for Registered Users</title>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.standard {
	font-family: Geneva, Arial, Helvetica, san-serif;
	font-size: 14px;
	font-style: normal;
	line-height: normal;
	font-weight: normal;
	font-variant: normal;
	text-transform: none;
	color: #333333;
	text-decoration: none;
}
.content {
	font-family: Geneva, Arial, Helvetica, san-serif;
	font-size: 14px;
	font-style: normal;
	line-height: normal;
	font-weight: normal;
	font-variant: normal;
	text-transform: none;
	color: #000000;
	text-decoration: none;
}
a {
	font-family: Geneva, Arial, Helvetica, san-serif;
	font-size: 14px;
	font-style: normal;
	line-height: normal;
	font-weight: normal;
	font-variant: normal;
	text-transform: none;
	color: #333333;
	text-decoration: none;
}
-->
</style>
</HEAD>
<BODY bgcolor="#666666">
Welcome <%= Session("Username") %>, to Wols Members Page.
<P>&nbsp;</P>
<a href="signup.asp">Signup/Remove to a LAN</a>
<br>
<a href="Register.asp?Update=TRUE">Update User Info</a>
<P></P>
<%
If Session("UserID") = 2 OR Session("UserID") = 1 then
	Response.Write "<hr width=130 align=left>"
	Response.Write "<strong><font size=+1>Admin Section</font></strong>"
	Response.Write "<BR><a href=createLAN.asp>Create a LAN</a>"
	Response.Write "<BR><a href=deleteLAN.asp>Delete a LAN</a>"
	Response.Write "<BR><a href=MailOut.asp>Send E-Mail to everyone</a>"
	Response.Write "<hr width=130 align=left>"
End If
%>

</body>
</html>