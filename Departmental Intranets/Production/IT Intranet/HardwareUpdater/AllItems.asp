<%@LAnguage = VBScript%>
<!-- #INCLUDE FILE="FormatDate.asp" -->
<%
	'ensure that this page is refreshed EVERY TIME
	Response.Expires = 0
	
	'define the XML object
	Dim XML
	Set XML = Session("XML")
	
	'if the data has been posted back, then the user has requested to save the file
	if Request.QueryString("PostBack") = "true" Then
		
		'check if the user has cancelled the changes
		If Len(Request.QueryString("Cancel")) > 0 Then
			'do very little...
		Else
			
			'rename the old Hardware.xml file
			Dim lstrBackupFileName
			Dim FS
			Set FS = Server.CreateObject("Scripting.FileSystemObject")
			
			lstrBackupFileName = Application("XMLFilePath") & FormatDate(Now, C_FormatDate_Identifier) & " Backup of Hardware.xml"
			
			on error resume next
			FS.MoveFile Application("XMLFilePath") & "Hardware.xml", lstrBackupFileName
			if FS.GetFile(lstrBackupFileName) is nothing then
				Response.Write "There was a problem backing up the previous version of the XML file. Please contact an administrator"
				Response.End
			end if
			on error goto 0
			
			'save the XML
			on error resume next
			XML.Save Application("XMLFilePath") & "Hardware.xml"
			if FS.GetFile(Application("XMLFilePath") & "Hardware.xml") is nothing then
				Response.Write "There was a problem saving the new XML file. Please contact an administrator"
				Response.End
			end if
			on error goto 0
		End If
		
		'abandon this session and redirect the user
		session.Abandon
		Response.Redirect "/"
	End If

'this subroutine will display all of the groups and items on the page
Sub ShowData()
		
	'the XML loaded successfully, so display the data as required
	Set lobjNodes = XML.selectNodes("HARDWARE/GROUP")
	For Each lvarTemp in lobjNodes
			
		'write the group information to the page
		If lvarTemp.selectSingleNode("@title").nodeValue <> "Other..." Then
			Response.Write "<SPAN class=Copyright><A href=""EditGroup.asp?GroupName=" & ReplaceBadChars(lvarTemp.selectSingleNode("@title").nodeValue) & """>" & lvarTemp.selectSingleNode("@title").nodeValue & "</A></SPAN><BR>"
			ShowChildren lvarTemp
		End If
	Next
End Sub

'this subroutine will display all of the items on the page
'
'INPUT:
'	pobjNode	- A single XML node object.
Sub ShowChildren(pobjNode)
	
	'define the local variables
	Dim lobjNodes
	Dim lvarTemp
	
	'loop through all of the items in the specified node
	Set lobjNodes = pobjNode.selectNodes("ITEM/@desc")
	For Each lvarTemp in lobjNodes
		
		'add the item to the page
		Response.Write "&nbsp;&nbsp;&nbsp;&nbsp;<SPAN class=Copyright><A href=""EditItem.asp?ItemName=" & ReplaceBadChars(lvarTemp.nodeValue) & """>" & lvarTemp.nodeValue & "</A></SPAN><BR>"
	Next
	Response.Write "<BR>"
End Sub

'this function will replace invalid characters in a name so it can be sent to
'the receiving page via the "GET" method
'
'INPUT:
'	TheString	- a string that needs to have bad characters replaced
'
'OUTPUT:
'	Strin		- A "GET" compliant version of "TheString"
Function ReplaceBadChars(TheString)
	
	'define the local variables
	Dim lstrTemp
	
	'replace the bad characters
	lstrTemp = TheString
	
	'lstrTemp = Replace(lstrTemp, "&amp;", "&")
	'lstrTemp = Replace(lstrTemp, "&quot;", """")
	
	lstrTemp = Replace(lstrTemp, "+", "%2B")
	lstrTemp = Replace(lstrTemp, "&", "%26")
	lstrTemp = Replace(lstrTemp, """", "%22")
	
	'return the fixed string
	ReplaceBadChars = lstrTemp
End Function
%>
<HTML>
<HEAD>
<TITLE>Update Hardware Catalogue</TITLE>
<LINK rel="stylesheet" type="text/css" href="../Styles/Style.css"></LINK>
</HEAD>

<BODY>
	<P class=MainHeading>Update Hardware Catalogue</P>
	<P class=SubHeading>All Items</P>
	
	<%ShowData()%>
	
	<FORM name=frmAllItems method=get>
		<INPUT type=hidden name=PostBack value=true>
		<INPUT type=Submit value="Save Changes" style="WIDTH:125px">&nbsp;<INPUT type=submit name=Cancel value="Cancel Changes" style="WIDTH:125px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=button onclick="window.open('EditGroup.asp', '_top')" value="Create New Group" style="WIDTH:125px">
	</FORM>
</BODY>
</HTML>