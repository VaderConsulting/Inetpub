<!--#include file="Database.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<title>Register</title>
	
<SCRIPT language="JavaScript">
<!--
  function VerifyData()
  {
    if (document.frmUser.Password.value != document.frmUser.VerifyPassword.value)
    {
      alert ("Your passwords do not match - please reenter");
      return false;
    }
    else 
		if (document.frmUser.Username.value == "")
		    {
		      alert ("Please enter a username");
		      return false;
		    }
		else
			if (document.frmUser.Password.value == "")
		    {
		      alert ("Please enter a password");
		      return false;
		    }
			else
		      return true;
  }
-->
</SCRIPT>
<!-- InstanceEndEditable --> 
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<style type="text/css">
<!--
.standard {
	font-family: Geneva, Arial, Helvetica, san-serif;
	font-size: 12px;
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
	font-size: 12px;
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
	font-size: 12px;
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

<% 

If Request("Username") = "False" Then
	Response.Write "<H3>Username Found!!!</H3>"
End If

%>
Please fill in ALL of the information including a valid e-mail address... WoL uses a script to go through the database to send e-mails to everyone
and also the website will send you e-mails if you are in a waiting list notifying you of changes...
<BR>
Your e-mail address will NEVER be given out to anyone without your conscent!!!
<form action="AddUser.asp" NAME="frmUser" METHOD="POST" onSubmit="return VerifyData()">
<input type="hidden" name="Update" value="<%= Request("Update") %>">
<table>
<tr>
	<td>First Name:</td>
	<td><input type="text" name="FirstName" value="<%= Session("FirstName") %>" size="40"></td>
</tr>
<tr>
	<td>Last Name:</td>
	<td><input type="text" name="LastName" value="<%= Session("LastName") %>" size="40"></td>
</tr>
<tr>
	<td>Gaming Alias:</td>
	<td><input type="text" name="Username" value="<%= Session("Username") %>" size="40"></td>
</tr>
<tr>
	<td>E-Mail Address:</td>
	<td><input type="text" name="EMailAddress" value="<%= Session("EMailAddress") %>" size="40"></td>
</tr>
<tr>
	<td>Password:</td>
	<td><input type="password" name="Password" value="<%= Session("Password") %>" size="40"></td>
</tr>
<tr>
	<td>Re-Type Password:</td>
	<td><input type="password" name="VerifyPassword" size="40"></td>
</tr>
<tr>
	<td><BR></td>
	<td><input type="submit" name="Submit" value="Register">&nbsp;&nbsp;<input type="Reset"></td>
</tr>
</table>
</form>

</body>
</html>
