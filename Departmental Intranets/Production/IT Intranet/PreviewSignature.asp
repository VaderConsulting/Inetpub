<% Option Explicit %>
<!-- #Include file="Includes/Menu.asp" -->
<%
Const mstrcHTMLTemplate = "D:\INETPUB\Departmental Intranets\Development\IT Intranet\EmailSig\HTMTemplate.htm.template"
Const mstrcFieldTemplate = "D:\INETPUB\Departmental Intranets\Development\IT Intranet\EmailSig\EMailSignatureConfig.xml"
Const mstrcFieldNameIndicatorSting = "##"

Dim mstrHTML
Dim strUserAccount
Dim strDisplayUser

strUserAccount = UCase(Request.ServerVariables("LOGON_USER"))
strDisplayUser = UCASE(Replace(strUserAccount, "OXIANA\", ""))

mstrHTML = GenerateEmailSignature(mstrcHTMLTemplate,mstrcFieldTemplate,mstrcFieldNameIndicatorSting,strDisplayUser)

Function GenerateEmailSignature(ByVal strTemplatePath, ByVal strFieldXMLPath, ByVal strFieldNameIndicator, ByVal strUserName)
  Dim strHTML
  Dim fsoFileSystem
  Dim fleHTML
  Dim dicProperties

  ' Prepopulate AD properties
  Set dicProperties = CreateObject("Scripting.Dictionary")

  Dim strFileContents

  Set fsoFileSystem = Server.CreateObject("Scripting.FileSystemObject")
  Set fleHTML = fsoFileSystem.OpenTextFile(strTemplatePath)
  strFileContents = fleHTML.ReadAll
  fleHTML.Close 

  Dim intOffset
  Dim intIndex1
  Dim intIndex2
  Dim strField
  Dim strValue


  intOffset = 1
  strHTML = ""
  Do
    intIndex1 = InStr(intOffset,strFileContents,strFieldNameIndicator)
    If (intIndex1 > 0) Then
      intIndex2 = InStr(intIndex1 + 1,strFileContents,strFieldNameIndicator)
      If (intIndex2 > 0) Then
      
        strField = Mid(strFileContents,intIndex1 + Len(strFieldNameIndicator),intIndex2 - intIndex1 - Len(strFieldNameIndicator))
        If Not dicProperties.Exists(strField) Then
          dicProperties.Add CStr(strField),CStr(GetField(strUserName,strField,strFieldXMLPath))
        End If
        strHTML = strHTML + Mid(strFileContents,intOffset,intIndex1 - intOffset) + dicProperties(strField)
        intOffset = intIndex2 + Len(strFieldNameIndicator)
      Else
        ' Naughty Naughty.  Non-closing 'strFieldNameIndicator'
        intIndex1 = -1
      End If
    End If
  Loop Until (intIndex1 = 0) Or (intOffset >= Len(strFileContents))

  If (intOffset < Len(strFileContents)) Then strHTML = strHTML + Mid(strFileContents,intOffset,Len(strFileContents) - intOffset)

  GenerateEmailSignature = strHTML
End Function

Function GetField(ByVal strUsername, ByVal strFieldName, ByVal strFieldXMLPath)
  Dim xmlDoc
  Dim xmlRootElement
  Dim xmlNodeList
  Dim xmlNode
  Dim strOutput
  Dim adoConn
  Dim adoRS
  Dim adoCmd
  Dim strADFieldValue
  Dim blnUserLowerCase
  Dim blnStripCRLF
  Dim strReplaceCRLFWith
  Dim strADFieldName
  Dim strHTMLPrefix
  
  strOutput = strFieldName
  
  ' Open XML file
  Set xmlDoc = Server.CreateObject("Microsoft.XMLDom")
  xmlDoc.async = False
  xmlDoc.load strFieldXMLPath
  Set xmlRootElement = xmlDoc.DocumentElement
  ' Find the field
  Set xmlNodeList = xmlDoc.SelectNodes("config/field[@name='" + strFieldName + "'][@adfieldname]")
  If (xmlNodeList.length = "0") Then GetField = "[Unknown Field """ + strFieldName + """]":Exit Function
  ' Get the field definition
  strADFieldName = GetXMLAttribute(xmlNodeList.item(0),"adfieldname","")
  blnUserLowerCase = (GetXMLAttribute(xmlNodeList.item(0),"lowercase","") <> "")
  blnStripCRLF = (GetXMLAttribute(xmlNodeList.item(0),"stripcrlf","") <> "")
  strReplaceCRLFWith = GetXMLAttribute(xmlNodeList.item(0),"replacecrlfwith","")
  ' Get the HTM Prefix
  Set xmlNodeList = xmlDoc.SelectNodes("config/field[@name='" + strFieldName + "'][@adfieldname]/prefix/htm")
  If (xmlNodeList.length = 1) Then
    Set xmlNode = xmlNodeList.item(0)
    strHTMLPrefix = xmlNode.text
  Else
    strHTMLPrefix = ""
  End If
  ' XML Cleanup
  Set xmlDoc = Nothing

  ' Sanity Checks
  If (strADFieldName = "") Then GetField = "[Field """ + strFieldName + """ has no AD FieldName]":Exit Function

  ' Query AD
  ' Create the connection object
	Set adoConn = CreateObject("ADODB.Connection")
	' Attempt to locate the user account in AD
	adoConn.Provider = "ADsDSOObject"
	adoConn.Properties("User ID") = CStr(Application("ADServiceAccount"))
	adoConn.Properties("Password") = CStr(Application("ADServicePassword"))
	adoConn.Properties("ADSI Flag") = 1
	adoConn.Properties("Encrypt Password") = True
	adoConn.Open "ADSI"
	Set adoCmd = CreateObject("ADODB.Command")
	Set adoCmd.ActiveConnection = adoConn
	adoCmd.CommandText = "<LDAP://oxiana.com.au/DC=Oxiana,DC=com,DC=au>;(&(objectCategory=user)(sAMAccountName=" + strUserName + "));" + strADFieldName + ";subtree"
	On Error Resume Next
	Set adoRS = adoCmd.Execute
	If Not IsNull(adoRS(strADFieldName)) then
	  strADFieldValue = CStr(adoRS(strADFieldName))
	Else
	  strADFieldValue = ""
	End If
	If (Err.Number <> 0) Then strADFieldValue = ""
	On Error Goto 0
	adoRS.Close

  ' Modify the text if specified
  If blnUserLowerCase Then strADFieldValue = LCase(strADFieldValue )
  If blnStripCRLF Then
    strADFieldValue = Replace(strADFieldValue,Chr(10),Chr(13))
    While InStr(strADFieldValue,Chr(13) + Chr(13)) > 0
      strADFieldValue = Replace(strADFieldValue,Chr(13) + Chr(13),Chr(13))
    Wend
    strADFieldValue = Replace(strADFieldValue,Chr(13),strReplaceCRLFWith)
  End If
  
  ' Remove the prefix if there is no value
  If (strADFieldValue = "") Then strHTMLPrefix = ""

  GetField = strHTMLPrefix + strADFieldValue
End Function

Function GetXMLAttribute(ByRef xmlNode, ByVal strAttributeName, ByVal strDefault)
  Dim intIndex
  GetXMLAttribute = strDefault
  For intIndex = 0 To xmlNode.attributes.length - 1
    If (UCase(xmlNode.attributes.Item(intIndex).Name) = UCase(strAttributeName)) Then GetXMLAttribute = xmlNode.attributes.Item(intIndex).Value
  Next
End Function

%>
<HTML>
<HEAD>
	<TITLE>IT Intranet: Preview My Email Signature</TITLE>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
	<SCRIPT language=JavaScript>
	<!--
	function HideNonPrintables()
	{
		for (Counter=0;Counter<HidePrint.length;Counter++) HidePrint[Counter].style.display = 'None';
	}
	function ShowNonPrintables()
	{
		for (Counter=0;Counter<HidePrint.length;Counter++) HidePrint[Counter].style.display = 'Inline';
	}
	//-->
	</SCRIPT>
</HEAD>

<BODY onbeforeprint="HideNonPrintables()" onafterprint="ShowNonPrintables()">
	<TABLE width=770 cellpadding=0 cellspacing=0 align=center width=770>
		<TR id=HidePrint height=79><TD colspan=3 background=Images/Logo.gif align=right class=Copyright><A href="SiteMap.asp">Site Map</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="ContactUs.asp">Contact Us</A></TD></TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint>
			<TD width=190><DIV class=HomeLink><A href="index.asp" width=100%>HOME</A></DIV></TD>
			<TD width=10>&nbsp;</TD>
			<TD width=570>&nbsp;</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR>
			<TD id=HidePrint width=190 valign=top>
				<TABLE cellpadding=0 cellspacing=0 width=100%>
					<%ShowMenu%>
				</TABLE>
			</TD>
			<TD width=10></TD>
			<TD width=570 valign=top>
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Preview My Email Signature</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
<table cellpadding="2" cellspacing="0" style="border:silver solid 1px;"><tr><td><%= mstrHTML %></td></tr></table>
<!--				<DIV class=SubHeading>Sub Heading</DIV>
				Content to be added... -->
				<DIV class=Copyright><BR>This page last updated: Sometime!! (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>