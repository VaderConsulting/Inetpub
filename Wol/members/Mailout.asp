<!--#include file="Database.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
<HEAD>
<TITLE>Mail to Everyone</TITLE>
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
<div align="center"><font size="+6">Mail to Everyone</font>
<form action="Mail2E1.asp" method="post">
<table>
<tr>
	<td>Subject:</td>
	<td><input type="text" name="Subject" size="66"></td>
</tr>
<tr>
	<td>Message Body:</td>
	<td><textarea cols="50" rows="12" name="body"></textarea></td>
</tr>
<tr>
	<td>&nbsp;</td>
	<td><input type="submit" name="Send" value="Send"></td>
</tr>
</table></div>




</form>

</body>
</html>
