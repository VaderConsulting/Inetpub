<%@LANGUAGE="VBSCRIPT"%> 
<!--#include file="Connections/WolLan.asp" -->
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
  strSQL = "SELECT * FROM lanInfo"
  Recordset1.Open strSQL, objConn
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/standard.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" -->
<title></title>
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
        <table width="100%" border="0" cellpadding="2" cellspacing="0" class="content">
          <% 
Do While Not Recordset1.EOF
%>
          <tr> 

            <td><%=(Recordset1.Fields.Item("lanName").Value)%></td>
            <td><div align="right">Date:<%=(Recordset1.Fields.Item("lanDate").Value)%> 
              </div></td>
          </tr>
          <tr> 

            <td colspan="2"><%=(Recordset1.Fields.Item("lanAddress").Value)%></td>
          </tr>
          <tr> 

            <td>Contact 
              <%=(Recordset1.Fields.Item("lanContact").Value)%> for more.</td>
            <td><div align="right"><%=(Recordset1.Fields.Item("lanTime").Value)%></div></td>
          </tr>
          <tr> 

            <td>Alcohol: 
              <%=(Recordset1.Fields.Item("lanAlchohol").Value)%> </td>
            <td><div align="right">Ages:<%=(Recordset1.Fields.Item("lanAges").Value)%></div></td>
          </tr>
          <tr> 

<td colspan="2"><div align="right">Lan 
                Cost: <%=(Recordset1.Fields.Item("lanCost").Value)%></div></td>
          </tr>
          <tr> 

            <td colspan="2">Extra 
              Notes: <%=(Recordset1.Fields.Item("lanNotes").Value)%></td>
          </tr>
          <% 
  Recordset1.MoveNext()
Loop
%>

<tr>
            <td class="content"><br>
              ..<br>
              &nbsp;</td>
          </tr>

        </table>

</body>
</html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>