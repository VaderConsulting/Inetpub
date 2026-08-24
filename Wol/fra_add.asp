
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/standard.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title>Sign Up</title>
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
<!-- InstanceBeginEditable name="head" --><!-- InstanceEndEditable --> 
</head>

<body bgcolor="#666666">


  
      <td bordercolor="#000000" bgcolor="#666666"> <!-- InstanceBeginEditable name="centerContent" -->
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
<tr>
            <td bgcolor="#333333" class="content"><font color="#CCCCCC"><strong>Sign 
              Up</strong></font></td>
          </tr>
          <tr>
            <td><span class="content">Please fill out the following form to register 
              for the upcoming LAN. </span> 
<form name="form1" method="POST" action="addtoDB.asp">

                <table width="100%" border="0" cellspacing="0" cellpadding="2">
                  
                  <tr valign="middle" class="content"> 
                    <td> 
<div align="right">Alias:</div></td>
                    <td> 
                      <input type="text" name="Alias" size="20"></td>
                  </tr>
                  <tr valign="middle" class="content"> 
<td> 
<div align="right">Email:</div></td>
                    <td> 
                      <input type="text" name="Email" size="50"></td>
                  </tr>
                  <tr valign="middle" class="content"> 
<td> 
<div align="right"></div></td>
                    <td> 
                      Your email address will not be displayed.</td>
                  </tr>
                  <tr valign="middle" class="content"> 
<td> 
<div align="right">Games:</div></td>
                    <td> 
                      <input type="text" name="Games" size="20"></td>
                  </tr>
                  <tr class="content"> 
                    <td> <div align="right"></div></td>
                    <td><input type="submit" name="Submit" value="Submit">
                    </td>
                  </tr>
                </table>
              
                <input type="hidden" name="MM_insert" value="form1">
              </form></td>
          </tr>
        </table>
</body>
</html>

<%
'Recordset1.Close()
'Set Recordset1 = Nothing
%>