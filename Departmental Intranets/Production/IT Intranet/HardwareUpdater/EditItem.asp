<%@LAnguage = VBScript%>

<%
	'ensure that this page is refreshed EVERY TIME
	Response.Expires = 0
	
	'define the local variables
	Dim XML
	Dim lobjNodes
	Dim lvarTemp
	Dim lvarTemp2
	Dim lobjItem
	Dim lblnFoundItem
	Set XML = Session("XML")
	
	'check if the user has posted data back to the form
	If Len(Request.QueryString("NewItemName")) > 0 Then
		
		'check if the user has requested to delete the current item
		If Len(Request.QueryString("Delete")) > 0 Then
			
			'delete the current item
			Session("TempNode").ParentNode.removeChild Session("TempNode")
			Response.Redirect "AllItems.asp"
		Else
			
			'check if the tempnode contains an object
			If Session("TempNode").NodeName = "GROUP" Then
				
				'define the local variables
				Dim lstrTemp
				Dim XML2
				
				'build the XML
				lstrTemp = ""
				lstrTemp = lstrTemp & "<ITEM desc=""" & ReplaceBadChars(Request.QueryString("NewItemName")) & """ cecode=""" & Request.QueryString("NewCECode") & """ oemcode=""" & Request.QueryString("NewOEMCode") & """ price=""" &  Request.QueryString("NewPrice") & """ mobilephone=""" & Request.QueryString("NewMobilePhone") & """ />"
				
				'create the XML object and load the new XML
				Set XML2 = Server.CreateObject("MSXML2.DOMDocument")
				XML2.loadXML(lstrTemp)
				
				'add the new node to the global XML document and return to all items
				Session("TempNode").InsertBefore XML2.selectSingleNode("ITEM"), FindNode("HARDWARE/GROUP/ITEM", "@desc", Request.QueryString("Node"))
				Response.Redirect "AllItems.asp"
				
			Else
				
				'Response.Write request.QueryString("NewItemName") & "<BR>"
				'Response.Write ReplaceBadChars(Request.QueryString("NewItemName"))
				'Response.End
				
				'update the selected item
				Session("TempNode").selectSingleNode("@desc").nodeValue = Request.QueryString("NewItemName")
				Session("TempNode").selectSingleNode("@cecode").nodeValue = Request.QueryString("NewCECode")
				Session("TempNode").selectSingleNode("@oemcode").nodeValue = Request.QueryString("NewOEMCode")
				Session("TempNode").selectSingleNode("@price").nodeValue = Request.QueryString("NewPrice")
				Session("TempNode").selectSingleNode("@mobilephone").nodeValue = Request.QueryString("NewMobilePhone")
				
				Session("TempNode").ParentNode.InsertBefore Session("TempNode"), FindNode("HARDWARE/GROUP/ITEM", "@desc", Request.QueryString("Node"))
				Response.Redirect "AllItems.asp"
			End If
		End If
	Else
		
		'loop through all of the items in the XML to locate the selected one
		Set lobjNodes = XML.SelectNodes("HARDWARE/GROUP/ITEM")
		For Each lvarTemp in lobjNodes
			
			'check if this is the selected item
			if lvarTemp.SelectSingleNode("@desc").nodeValue = Request.QueryString("ItemName") Then
				
				'this is the selected item, so keep a reference to it!
				Set lobjItem = lvarTemp
				Set Session("TempNode") = lobjItem
				lblnFoundItem = True
				Exit For
			End If
		Next
		Set lobjNodes = Nothing
	End If
	On Error Resume Next
%>
<HTML>
<HEAD>
<TITLE>Update Hardware Catalogue (Edit Item)</TITLE>
<LINK rel="stylesheet" type="text/css" href="../Styles/Style.css"></LINK>
</HEAD>

<BODY>
	<FORM name=frmEditItem method=get>
		<P class=MainHeading>Update Hardware Catalogue</P>
		<P class=SubHeading>Item Details</P>
		<SPAN style="WIDTH:150px">Name:</SPAN><INPUT name=NewItemName value="<%= replacebadchars(Request.QueryString("ItemName"))%>" style="WIDTH:400px"><BR>
		<SPAN style="WIDTH:150px">CE Code:</SPAN><INPUT name=NewCECODE value="<%= ReplaceBadChars(lobjItem.SelectSingleNode("@cecode").nodeValue)%>" style="WIDTH:100px"><BR>
		<SPAN style="WIDTH:150px">OEM Code:</SPAN><INPUT name=NewOEMCODE value="<%= ReplaceBadChars(lobjItem.SelectSingleNode("@oemcode").nodeValue)%>" style="WIDTH:100px"><BR>
		<SPAN style="WIDTH:150px">Price:</SPAN><INPUT name=NewPrice value="<%= ReplaceBadChars(lobjItem.SelectSingleNode("@price").nodeValue)%>" style="WIDTH:100px"><BR>
		<SPAN style="WIDTH:150px">Special Device Type:</SPAN><SELECT name=NewMobilePhone>
			<OPTION value="">-- Select --</OPTION>
			<OPTION value="Blackberry" <%If lobjItem.SelectSingleNode("@mobilephone").nodeValue = "Blackberry" Then Response.Write "Selected" End If%>>Blackberry</OPTION>
			<OPTION value="MobilePhone" <%If lobjItem.SelectSingleNode("@mobilephone").nodeValue = "MobilePhone" Then Response.Write "Selected" End If%>>Mobile Phone</OPTION>
		</SELECT><BR><BR>
		<SPAN style="WIDTH:150px">Move Item Above:</SPAN><SELECT name="Node">
			<OPTION value="">-- Select --</OPTION>
			<%
				Set lvarTemp = Nothing
				For Each lvarTemp In Session("XML").SelectNodes("HARDWARE/GROUP")
					
					'only display the nodes that are in the same group as the selected item
					If lvarTemp.SelectSingleNode("@title").nodeValue = Session("TempNode").ParentNode.SelectSingleNode("@title").nodeValue Then
					
						'loop through all of the items
						for each lvarTemp2 in lvarTemp.SelectNodes("ITEM")
							
							'check if this item should be displayed
							If lvarTemp2.SelectSingleNode("@desc").nodeValue <> Request.QueryString("ItemName") Then
								
								'display this item
								Response.Write "<OPTION value=""" & lvarTemp2.SelectSingleNode("@desc").nodeValue & """>" & lvarTemp2.SelectSingleNode("@desc").nodeValue & "</OPTION>"
							End If
						Next
						exit for
					End If
				Next
			%>
		</SELECT>
		<BR>
		<BR>
		<%IF lblnFoundItem = True Then%>
		<INPUT type=submit value="Save Changes" style="WIDTH:125px">&nbsp;<INPUT type=button value="Cancel Changes" onclick="window.open('AllItems.asp', '_top')" style="WIDTH:125px">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=submit name=Delete value="Delete Item" style="WIDTH:125px">
		<%Else%>
		<INPUT type=submit value="Create Item" style="WIDTH:125px">&nbsp;<INPUT type=button value="Cancel" onclick="window.open('AllItems.asp', '_top')" style="WIDTH:125px"><BR>
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
	lstrTemp = Replace(lstrTemp, "&amp;", "&")
	lstrTemp = Replace(lstrTemp, "&quot;", """")
	
	lstrTemp = Replace(lstrTemp, "&", "&amp;", 1)
	lstrTemp = Replace(lstrTemp, """", "&quot;", 1)
	'lstrTemp = Replace(lstrTemp, "'", "&apos;", 1)
	
	ReplaceBadChars = lstrTemp
End Function
%>