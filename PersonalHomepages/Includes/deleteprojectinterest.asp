<%@ Language=VBScript %>
<%'ABOVE LINE CAN BE REMOVED IF A SESSION IS REQUIRED
Option Explicit 
Response.Expires = true
'***************************************************************************
' Name:		deleteprojectinterest.asp
' Purpose:	Deletes project or interest
' Copyright (c)2001 Tusk Technologies. All rights reserved.
'
' Name		Date		Desc
' 
' 
' OUTSTANDING TASKS
' 
'**************************************************************************** 
'ENSURE IIS HAS BUFFERING TURNED ON OTHERWISE UNCOMMENT THIS LINE
'Response.Buffer = True

'UNCOMMENT THIS LINE WHEN BUSINESS LOGIC IS COMPLETE
On Error Resume Next

'STANDARD INCLUDE FILES
%>
<!--#INCLUDE FILE="constants.asp"-->
<!--#include file="whologgedin.asp"-->
<%
'***************DECLARE GLOBAL VARIABLES******************
Dim con
Dim com
Dim strWhichEntry
Dim strWhichNumber
Dim strAllString
Dim strNewString
Dim arystrSplit
'**********************FUNCTIONS**************************

sub fillArray(strType)
	' Uses the split function to fill an array
	if lcase(strType) = "projects" then
		arystrSplit = split(strAllString,"|")
	elseif lcase(strType) = "interests" then
		arystrSplit = split(strAllString,"|")
	end if
end sub

'*************************MAIN****************************
' Retrieve any query string variables
strWhichEntry = Request.QueryString("whichEntry")
strWhichNumber = Request.QueryString("whichNumber")
strAllString = Request.QueryString("allString")

set con = server.CreateObject("ADODB.Connection")
set com = server.CreateObject("ADODB.Command")
strNewString = ""

' Create the new string minus the deleted item
if strWhichEntry = "Project" then
	' First we need to split the array
	fillArray("Projects")
	' Now we need to rebuild the string without the deleted item
	for i = 0 to ubound(arystrSplit) step 1
		' Skip if deleted item
		if i <> cint(strWhichNumber) then
			' Else add to string
			if strNewString = "" then
				strNewString = arystrSplit(i)
			else
				strNewString = strNewString & "|" & arystrSplit(i)
			end if
		end if
	next
elseif strWhichEntry = "Interest" then
	' First we need to split the array
	fillArray("Interests")
	' Now we need to rebuild the string without the deleted item
	for i = 0 to ubound(arystrSplit) step 2
		' Skip if deleted item
		if i <> cint(strWhichNumber) then
			if strNewString = "" then
				strNewString = arystrSplit(i) & "|" & arystrSplit(i+1)
			else
				strNewString = strNewString & "|" & arystrSplit(i) & "|" & arystrSplit(i+1)
			end if
		end if
	next
end if

if false then
	Response.Write strWhichEntry & "<br><br>"
	Response.Write strWhichNumber & "<br><br>"
	Response.Write strAllString & "<br><br>"
	Response.Write strNewString & "<br><br>"
	Response.Write strSQL & "<br><br>"
	Response.End
else
	con.Open CONST_CONNECTION_STRING
	com.ActiveConnection = con
	com.CommandType = adCmdStoredProc
	com.Parameters.Append com.CreateParameter("@EmpNum", adChar, adParamInput, 10, UserNameReturn)
	
	if strWhichEntry = "Project" then
		com.CommandText = "spUpdateProjects"
		com.Parameters.Append com.CreateParameter("@strProjects", adVarChar, adParamInput, 5000, Trim(strNewString))
	elseif strWhichEntry = "Interest" then
		com.CommandText = "spUpdateInterests"
		com.Parameters.Append com.CreateParameter("@strInterests", adVarChar, adParamInput, 5000, Trim(strNewString))
	end if
	
	com.Execute
end if

if com.State <> 0 then com.Close
set com = nothing
if con.State <> 0 then con.Close
set con = nothing

if Err.number = 0 then
	Response.Redirect "edithomepage.asp?Action=Delete&Result=True"
else
	Response.Redirect "edithomepage.asp?Action=Delete&Result=False&ErrNum=" & Err.number & "&ErrDesc" & Err.Description
end if
%>