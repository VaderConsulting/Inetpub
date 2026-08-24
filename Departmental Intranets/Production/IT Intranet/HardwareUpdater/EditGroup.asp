<%@LAnguage = VBScript%>

<%
	'ensure that this page is refreshed EVERY TIME
	Response.Expires = 0
	
	'define the local variables
	Dim XML
	Dim lobjNodes
	Dim lvarTemp
	Dim lobjGroup
	Dim lblnFoundGroup
	Set XML = Session("XML")
	
	'check if the user has posted data back to the form
	If Len(Request.QueryString("NewGroupName")) > 0 Then
		
		'check if the user has requested to delete this group
		If Len(Request.QueryString("Delete")) > 0 Then
			
			'delete the current group
			Session("TempNode").ParentNode.removeChild Session("TempNode")
			Response.Redirect "AllItems.asp"
		
		'check if the user has requested to create an item
		ElseIf Len(Request.QueryString("Create")) > 0 Then
			
			'create a new item
			Response.Redirect "EditItem.asp"
			
		Else
			
			'check if this is a new group being added
			If Session("TempNode") Is Nothing Then
				
				'the user is adding a new group
				Dim lstrXML
				Dim XML2
				
				'create the XML for the new group
				lstrXML = "<GROUP title=""" & ReplaceBadChars(Request.QueryString("NewGroupName")) & """ />"
				
				'create the XML object and load the new XML
				Set XML2 = Server.CreateObject("MSXML2.DOMDocument")
				XML2.loadXML(lstrXML)
				
				'add the new group to the global XML document
				Session("XML").SelectSingleNode("HARDWARE").InsertBefore XML2.selectSingleNode("GROUP"), FindNode("HARDWARE/GROUP", "@title", Request.QueryString("Node"))
				Response.Redirect "AllItems.asp"
			Else
				
				'update the current group
				Session("TempNode").selectSingleNode("@title").nodeValue = ReplaceBadChars(Request.QueryString("NewGroupName"))
				Session("XML").SelectSingleNode("HARDWARE").InsertBefore Session("TempNode"), FindNode("HARDWARE/GROUP", "@title", Request.QueryString("Node"))
				Response.Redirect "AllItems.asp"
			End If
		End If
	Else
		
		Set lobjGroup = Nothing
		Set Session("TempNode") = Nothing
		
		'loop through all of the items in the XML to locate the selected one
		Set lobjNodes = XML.SelectNodes("HARDWARE/GROUP")
		For Each lvarTemp in lobjNodes
			
			'check if this is the selected item
			if lvarTemp.SelectSingleNode("@title").nodeValue = Request.QueryString("GroupName") Then
				
				'this is the selected item, so keep a reference to it!
				Set lobjGroup = lvarTemp
				Set Session("TempNode") = lobjGroup
				lblnFoundGroup = True
				Exit For
			End If
		Next
		Set lobjNodes = Nothing
	End If
	
	On Error Resume Next
%>
<HTML>
<HEAD>
<TITLE>Update Hardware Catalogue (Edit Group)</TITLE>
<LINK rel="stylesheet" type="text/css" href="../Styles/Style.css"></LINK>
</HEAD>

<BODY>
	<FORM name=frmEditItem method=get>
		<P class=MainHeading>Update Hardware Catalogue</P>
		<P class=SubHeading>Group Details:</P>
		<SPAN style="WIDTH:150px">Name:</SPAN><INPUT name=NewGroupName value="<%= lobjGroup.SelectSingleNode("@title").nodeValue%>" style="WIDTH:400px"><BR><BR>
		<SPAN style="WIDTH:150px">Move Group Above:</SPAN><SELECT name="Node">
			<OPTION value="">-- Select --</OPTION>
			<%
				For Each lvarTemp In Session("XML").SelectNodes("HARDWARE/GROUP")
					If lvarTemp.SelectSingleNode("@title").nodeValue <> Request.QueryString("GroupName") Then
						Response.Write "<OPTION value=""" & lvarTemp.SelectSingleNode("@title").nodeValue & """>" & lvarTemp.SelectSingleNode("@title").nodeValue & "</OPTION>"
					End If
				Next
			%>
		</SELECT>
		<BR>
		<BR>
		<%If lblnFoundGroup = True Then%>
		<INPUT type=submit value="Save Changes" style="WIDTH:125px">&nbsp;<INPUT type=button value="Cancel" onclick="window.open('AllItems.asp', '_top')" style="WIDTH:125px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=submit name=Delete value="Delete Group" style="WIDTH:125px">&nbsp;<INPUT type=submit name="Create" value="Create Item" style="WIDTH:125px">
		<%Else%>
		<INPUT type=submit value="Create Group" style="WIDTH:125px">&nbsp;<INPUT type=button value="Cancel" onclick="window.open('AllItems.asp', '_top')" style="WIDTH:125px">
		<%End If%>
	</FORM>
</BODY>
</HTML>
<%
'this function will search the XML for the specified parameters and 
'return a matching node if located.
'
'INPUT:
'	Path	- A string path to the nodelist that should be searched (ie "HARDWARE/GROUP")
'	attrib	- A string value that identifies the attribute to compare with the value (ie "@title")
'	value	- A string value that should be located in the specified attribute field
'
'OUTPUT:
'	MSXML2.IXMLDOMNode	- The node that matched the specified parameters, or "Nothing" is not matched
Function FindNode(Path, attrib, value)
	
	'define the local variables
	Dim lvarTemp
	
	'loop through all of the nodes
	For Each lvarTemp In Session("XML").SelectNodes(Path)
		
		'check if the specified attribute matches the specified value
		if lvarTemp.SelectSingleNode(attrib).nodevalue = value then
			
			'return this node to the calling function
			Set FindNode = lvarTemp
			exit Function
		end if
	Next
End Function

Function ReplaceBadChars(TheString)
	Dim lstrTemp
	
	lstrTemp = TheString
	lstrTemp = Replace(lstrTemp, "&", "&amp;")
	lstrTemp = Replace(lstrTemp, """", "&quot;")
	lstrTemp = Replace(lstrTemp, "'", "&apos;")
	
	ReplaceBadChars = lstrTemp
End Function
%>