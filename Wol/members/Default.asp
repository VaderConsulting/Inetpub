<%
If Not Session("UserID") = "" Then
	Response.Redirect "menu_for_registered_users.asp"
End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ISO-8859-1">
<META NAME="Generator" CONTENT="Microsoft FrontPage 5.0">
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

<FORM ACTION="CheckLogin.asp<% If Request("SecondTry")="True" Then                                  Response.Write "?SecondTry=True"                               End If %>" METHOD="POST">
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

<A HREF=Forgot.asp><font color="#FFFFFF">Forgot Password</font></a><P>
<A HREF=Register.asp><font color="#FFFFFF">Register A New Member</font></a><BR>
<BR>
<font color="#FFFFFF">If you are having Problems Signing up try this Link ---
</font> <a href="http://wol.awoodman.net/members/Default.asp" target="_blank">
            <font size="3" color="#FFFFFF">http://wol.awoodman.net/members/</font></a><font size="3"><br>
<P>
NOTE: <font face="Arial" size="2">If you are using Internet Explorer, goto 
Tools, Internet Options, Privacy, Edit... , Address of Web site
<a href="http://wol.awoodman.net">http://wol.awoodman.net</a> </font>then click 
on Allow, the Ok.<P>
under IE go to Tools &gt; Internet Options &gt; Privacy and change the cookies to low 
this works.<P>
Or<P>
Make sure your Internet Explorer Security settings are set to medium or lower, otherwise you will get errors...
<BR>I am looking into how to solve this issue...
</BODY>
</HTML>