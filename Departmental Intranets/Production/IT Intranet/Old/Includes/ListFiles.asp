<%
'define the constants
Dim ROOTPATH
ROOTPATH = Application("Documents")	'This constant will define the top level folder where the documents will be stored

'this subroutine will add all of the located documents to the page
Sub ShowDocs(SubFolder)
	
	'define the local variables
	Dim FS
	Dim Folder
	Dim lvarTemp
	Dim FileType
	Dim lstrFolderURL
	lstrFolderURL = Request.ServerVariables("URL") & "?SelectedFolder=" & Request.QueryString("SelectedFolder")
	if right(lstrFolderURL, 1) <> "/" Then
		lstrFolderURL = lstrFolderURL & "/"
	End If
	Dim lstrSelectedFolder
	lstrSelectedFolder = Request.QueryString("SelectedFolder")
	Dim larrFolders
	Dim llngCounter
	
	'check which folder is being displayed
	If len(lstrSelectedFolder) = 0 Then
		
		'this is currently the root folder, so dont display a link
		Response.Write "<P class=Copyright style=""TEXT-ALIGN:Left"">Root</P>"
	Else
		
		'this is not the root folder for this page, so display a link back to the root
		Response.Write "<P class=Copyright style=""TEXT-ALIGN:Left""><A href=""" & Request.ServerVariables("URL") & """>Root</A>"
		
		'loop through all of the user-selected subfolders and display them as links
		larrFolders = Split(lstrSelectedFolder, "/")
		For llngCounter = 1 To UBound(larrFolders) - 1
			Response.Write " \ <A href=""" & left(lstrFolderURL, instr(1, lstrFolderURL, larrFolders(llngCounter)) - 1) & larrFolders(llngCounter) & """>" & larrFolders(llngCounter) & "</A>"
		Next
		
		'display the currently selected folder without a link.
		Response.Write " \ " & larrFolders(llngCounter) & "</P>"
	End If
	
	'locate the requested folder
	Set FS = Server.CreateObject("Scripting.FileSystemObject")
	
	On Error Resume Next
	Set Folder = FS.GetFolder(Server.MapPath(ROOTPATH) & "\" & Replace(SubFolder & lstrSelectedFolder, "/", "\"))
	On Error Goto 0
	
	'check that the folder exists
	If Typename(Folder) = "Empty" Then
		
		'inform the user that the folder did not exist
		Response.Write "<SPAN class=Warning>The specified folder '" & ROOTPATH & Replace(SubFolder, "\", "/") & lstrSelectedFolder & "' does not exist. Please contact an administrator</SPAN>"
	
	'check if any files exist in the specified folder
	ElseIf (Folder.Files.Count = 0) and (folder.subfolders.count = 0) then
		
		'inform the user that there were no files
		Response.Write "<SPAN class=Warning>This folder contains no files or folders.</SPAN>"
	Else
		
		'loop through all of the subfolders
		For Each lvarTemp in Folder.SubFolders
			
			'add each subfolder to the page
			Response.Write "<A href=""" & lstrFolderURL & ReplaceInvalidChars(lvarTemp.name) & """><IMG height=20 width=20 src=""Images/Folder.gif"" border=0 align=""absmiddle"">&nbsp;" & lvarTemp.name & "</A><BR>"
		Next
		
		'loop through all of the files and display them
		For Each lvarTemp In Folder.Files
			
			'determine the image to display based on the type of file
			Select Case lvarTemp.Type
				Case "Adobe Acrobat Document", "Adobe Acrobat 7.0 Document": FileType = "PDF"
				Case "Microsoft Excel Worksheet": FileType = "Excel"
				Case "Microsoft Word Document": FileType = "Word"
				Case "Text Document": FileType = "Text"
				Case Else: FileType = lvarTemp.Type
			End Select
			
			'display a link to the located file
			Response.Write "<A href=""" & ROOTPATH & SubFolder & lstrSelectedFolder & "/" & lvarTemp.name & """><IMG height=20 width=20 src=""Images/" & FileType & ".gif"" border=0 align=""absmiddle"">&nbsp;" & mid(lvarTemp.name, 1, instr(1,lvarTemp.name, ".") - 1) & "</A><BR>"
		Next
	End If
	
	'release the filesystem objects
	Set Folder = Nothing
	Set FS = Nothing
End Sub

'this function will replace the invalid characters in the OriginalValue parameter
'and will return the HTML friendly representation
'
'INPUT:
'	OriginalValue	- A String value that is to be formatted for HTML compatibility
'
'OUTPUT:
'	String			- The formatted version of the "OriginalValue" Parameter
Private Function ReplaceInvalidChars(OriginalValue)
	
	'define the local variables
	Dim lstrTemp
	lstrTemp = OriginalValue
	
	'replace the invalid characters that web pages wont display
	lstrTemp = Replace(lstrTemp, "%", "%25")
	lstrTemp = Replace(lstrTemp, "+", "%2b")
	lstrTemp = Replace(lstrTemp, "&", "%26")
	lstrTemp = Replace(lstrTemp, "#", "%23")
	lstrTemp = Replace(lstrTemp, ",", "%2c")
	
	'return the formatted value
	ReplaceInvalidChars = lstrTemp
End Function
%>