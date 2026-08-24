<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/WolLan.asp" -->
<%
' *** Restrict Access To Page: Grant or deny access to this page
MM_authorizedUsers=""
MM_authFailedURL="../index.asp"
MM_grantAccess=false
If Session("MM_Username") <> "" Then
  If (true Or CStr(Session("MM_UserAuthorization"))="") Or _
         (InStr(1,MM_authorizedUsers,Session("MM_UserAuthorization"))>=1) Then
    MM_grantAccess = true
  End If
End If
If Not MM_grantAccess Then
  MM_qsChar = "?"
  If (InStr(1,MM_authFailedURL,"?") >= 1) Then MM_qsChar = "&"
  MM_referrer = Request.ServerVariables("URL")
  if (Len(Request.QueryString()) > 0) Then MM_referrer = MM_referrer & "?" & Request.QueryString()
  MM_authFailedURL = MM_authFailedURL & MM_qsChar & "accessdenied=" & Server.URLEncode(MM_referrer)
  Response.Redirect(MM_authFailedURL)
End If
%>
<%
' *** Edit Operations: declare variables

Dim MM_editAction
Dim MM_abortEdit
Dim MM_editQuery
Dim MM_editCmd

Dim MM_editConnection
Dim MM_editTable
Dim MM_editRedirectUrl
Dim MM_editColumn
Dim MM_recordId

Dim MM_fieldsStr
Dim MM_columnsStr
Dim MM_fields
Dim MM_columns
Dim MM_typeArray
Dim MM_formVal
Dim MM_delim
Dim MM_altVal
Dim MM_emptyVal
Dim MM_i

MM_editAction = CStr(Request.ServerVariables("SCRIPT_NAME"))
If (Request.QueryString <> "") Then
  MM_editAction = MM_editAction & "?" & Request.QueryString
End If

' boolean to abort record edit
MM_abortEdit = false

' query string to execute
MM_editQuery = ""
%>
<%
' *** Delete Record: declare variables

if (CStr(Request("MM_delete")) = "form1" And CStr(Request("MM_recordId")) <> "") Then

  MM_editConnection = MM_WolLan_STRING
  MM_editTable = "lanInfo"
  MM_editColumn = "ID"
  MM_recordId = "" + Request.Form("MM_recordId") + ""
  MM_editRedirectUrl = "../index.asp"

  ' append the query string to the redirect URL
  If (MM_editRedirectUrl <> "" And Request.QueryString <> "") Then
    If (InStr(1, MM_editRedirectUrl, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
      MM_editRedirectUrl = MM_editRedirectUrl & "?" & Request.QueryString
    Else
      MM_editRedirectUrl = MM_editRedirectUrl & "&" & Request.QueryString
    End If
  End If
  
End If
%>
<%
' *** Delete Record: construct a sql delete statement and execute it

If (CStr(Request("MM_delete")) <> "" And CStr(Request("MM_recordId")) <> "") Then

  ' create the sql delete statement
  MM_editQuery = "delete from " & MM_editTable & " where " & MM_editColumn & " = " & MM_recordId

  If (Not MM_abortEdit) Then
    ' execute the delete
    Set MM_editCmd = Server.CreateObject("ADODB.Command")
    MM_editCmd.ActiveConnection = MM_editConnection
    MM_editCmd.CommandText = MM_editQuery
    MM_editCmd.Execute
    MM_editCmd.ActiveConnection.Close

    If (MM_editRedirectUrl <> "") Then
      Response.Redirect(MM_editRedirectUrl)
    End If
  End If

End If
%>
<%
Dim Recordset1__MMColParam
Recordset1__MMColParam = "1"
If (Request.QueryString("ID") <> "") Then 
  Recordset1__MMColParam = Request.QueryString("ID")
End If
%>
<%
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_WolLan_STRING
Recordset1.Source = "SELECT * FROM lanInfo WHERE ID = " + Replace(Recordset1__MMColParam, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html><!-- InstanceBegin template="/Templates/standard.dwt.asp" codeOutsideHTMLIsLocked="false" -->
<head>
<!-- InstanceBeginEditable name="doctitle" --> 
<title>Untitled Document</title>
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

<body bgcolor="#CCCCCC">
<div align="center"> 
  <p class="standard">&nbsp;</p>
  <table width="90%" border="1" cellpadding="2" cellspacing="0" bordercolor="#000000">
    <tr> 
      <td colspan="3" bordercolor="#000000"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td bgcolor="#CCCCCC"><img src="../Images/title.gif" width="209" height="71"></td>
            <td bgcolor="#CCCCCC"> <div align="right"><img src="../Images/image.jpg" width="76" height="71"></div></td>
          </tr>
        </table></td>
    </tr>
    <tr align="left" valign="top"> 
      <td width="150" bordercolor="#000000" bgcolor="#999999" class="standard"><p><br>
          <strong>Navigation</strong><br>
          <img src="../images/spacer.gif" width="5" height="300" border="0" align="left"> 
          <a href="../laninfo.asp">Lan Info</a><br>
          <a href="../add.asp">Signup</a><br>
          <a href="../view.asp">View Enteries</a><br>
          <a href="../tobring.asp">What to Bring</a><br>
          <a href="../map.asp">Map</a><br>
          <br>
          <a href="../thanks.asp">Thanks to...</a></p>
        </td>
      <td bordercolor="#000000" bgcolor="#666666"> <!-- InstanceBeginEditable name="centerContent" --> 
        <form name="form1" method="POST" action="<%=MM_editAction%>">

          <table width="100%" border="0" cellspacing="0" cellpadding="2">
            
            <tr class="content"> 
              <td> 
<div align="right">LAN Number:</div></td>
              <td> 
                <div align="left"><%=(Recordset1.Fields.Item("ID").Value)%></div></td>
            </tr>
            <tr class="content"> 
<td> 
<div align="right">LAN Name:</div></td>
              <td> 
                <div align="left"><%=(Recordset1.Fields.Item("lanName").Value)%></div></td>
            </tr>
            <tr class="content"> 
<td> 
<div align="right">LAN Date:</div></td>
              <td> 
                <div align="left"><%=(Recordset1.Fields.Item("lanDate").Value)%></div></td>
            </tr>
            <tr class="content"> 
<td> 
<div align="right"></div></td>
              <td> 
                <div align="left">
                  <input type="submit" name="Submit" value="Submit">
                </div></td>
            </tr>
          </table>
        
          <input type="hidden" name="MM_delete" value="form1">
          <input type="hidden" name="MM_recordId" value="<%= Recordset1.Fields.Item("ID").Value %>">
        </form>
        <!-- InstanceEndEditable --></td>
      <td width="150" bordercolor="#000000" bgcolor="#999999" class="standard"><strong><br>
        Next Lan :</strong><br>
        <img src="../images/spacer.gif" width="5" height="300" border="0" align="left">Free 
        LAN<br>
        15th June<br>
        Lot 2 Gnangara Drive </td>
    </tr>
    <tr align="left" valign="top" bgcolor="#333333"> 
<td colspan="3" bordercolor="#000000"> 
<div align="center" class="standard"><font color="#CCCCCC">&pound;ogaan.net<font size="1"> 
          &copy;</font></font></div></td>
    </tr>
  </table>
</div>
</body>
<!-- InstanceEnd --></html>
<%
Recordset1.Close()
Set Recordset1 = Nothing
%>
