<%
If Not Session("UserID") = "" Then
	Response.Redirect "menu_for_registered_users.asp"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<META NAME="Generator" CONTENT="NetObjects Fusion 5.0.1 for Windows">
<TITLE>Login</TITLE>
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
  If Request("SecondTry") = "True" Then        ' User's second attempt at logging in
    If Request("WrongPW") = "True" Then 
      Response.Write "<font color=ff0000>Invalid Password!</font> Please try again: " & _
                     "if you get it wrong we'll ask you to re-register.<BR>"
    Else
      Response.Write "<font color=ff0000>Alias not found. Please try again.</font> " & _
                     "<BR>"
    End If
  End If
  Response.Write "Please enter your Alias and Password " & _
                 "to login to the system."
%>

<FORM ACTION="CheckLogin.asp<% If Request("SecondTry")="True" Then 
                                 Response.Write "?SecondTry=True"
                               End If %>" METHOD="POST">
  <TABLE BORDER=0>
    <TR>
      <TD>Gaming Alias:</TD>
      <TD><INPUT TYPE="Text" NAME="Username" 
          <% If Request("SecondTry") = "True" then %> 
            VALUE="<%= Session("Username") %>"
          <% End If %> 
          SIZE="40"></TD>
    </TR>
    <TR>
      <TD>Password:</TD>
      <TD><INPUT TYPE="Password" NAME="Password" SIZE="40"></TD>
    </TR>
    <TR>
      <TD></TD>
      <TD align=center><INPUT TYPE="Submit" VALUE="Login">&nbsp;&nbsp;
                       <INPUT TYPE="RESET"></TD>
    </TR>
  </TABLE>
</FORM>

<A HREF=Forgot.asp>Forgot Password</a><P>
<A HREF=Register.asp>Register</a><BR>

</BODY>
</HTML>
 
