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
Dim Recordset1
Dim Recordset1_numRows

Set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_WolLan_STRING
Recordset1.Source = "SELECT * FROM lanInfo"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 1
Recordset1.Open()

Recordset1_numRows = 0
%>
<%
Dim Repeat1__numRows
Dim Repeat1__index

Repeat1__numRows = -1
Repeat1__index = 0
Recordset1_numRows = Recordset1_numRows + Repeat1__numRows
%>
<%
Dim MM_paramName 
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

Dim MM_keepNone
Dim MM_keepURL
Dim MM_keepForm
Dim MM_keepBoth

Dim MM_removeList
Dim MM_item
Dim MM_nextItem

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then
  MM_removeList = MM_removeList & "&" & MM_paramName & "="
End If

MM_keepURL=""
MM_keepForm=""
MM_keepBoth=""
MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each MM_item In Request.QueryString
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & MM_nextItem & Server.URLencode(Request.QueryString(MM_item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each MM_item In Request.Form
  MM_nextItem = "&" & MM_item & "="
  If (InStr(1,MM_removeList,MM_nextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & MM_nextItem & Server.URLencode(Request.Form(MM_item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
If (MM_keepBoth <> "") Then 
  MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
End If
If (MM_keepURL <> "")  Then
  MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
End If
If (MM_keepForm <> "") Then
  MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)
End If

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
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
        <% 
While ((Repeat1__numRows <> 0) AND (NOT Recordset1.EOF)) 
%>
        <span class="content">Please select LAN to remove: </span> 
        <table width="100%" border="0" cellspacing="0" cellpadding="2">
          <tr> 
            <td width="25%" class="content"><%=(Recordset1.Fields.Item("ID").Value)%></td>
            <td width="25%" class="content"><%=(Recordset1.Fields.Item("lanName").Value)%></td>
            <td width="25%" class="content"><%=(Recordset1.Fields.Item("lanDate").Value)%></td>
            <td width="25%" class="content"><A HREF="confirm.asp?<%= MM_keepBoth & MM_joinChar(MM_keepBoth) & "ID=" & Recordset1.Fields.Item("ID").Value %>">Remove</A></td>
          </tr>
        </table>
        <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  Recordset1.MoveNext()
Wend
%>
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
