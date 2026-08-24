<%@ Language=VBScript %>
<%'ABOVE LINE CAN BE REMOVED IF A SESSION IS REQUIRED
Option Explicit 
Response.Expires = true
'***************************************************************************
' Name:		saveprojectinterest.asp
' Purpose:	Saves fields
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
'On Error Resume Next

'STANDARD INCLUDE FILES
%>
<!--#INCLUDE FILE="constants.asp"-->
<!--#include file="whologgedin.asp"-->
<%
'***************DECLARE GLOBAL VARIABLES******************
Dim con
Dim com
Dim strWhichEntry
Dim strString
Dim strPhoneNumber, strRoomNumber, strAltContact
'**********************FUNCTIONS**************************

'*************************MAIN****************************
' Retrieve any query string variables
strWhichEntry = Request.QueryString("whichEntry")
strString = Request.QueryString("newEntry")
strPhoneNumber = Request.QueryString("PhoneNumber")
strRoomNumber = Request.QueryString("RoomNumber")
strAltContact = Request.QueryString("AltContact")

set con = server.CreateObject("ADODB.Connection")
set com = server.CreateObject("ADODB.Command")

if false then
	Response.Write strWhichEntry & "<br><br>"
	Response.Write strString & "<br><br>"
	Response.End
else
	
	con.Open CONST_CONNECTION_STRING
	com.ActiveConnection = con
	com.CommandType = adCmdStoredProc
	com.Parameters.Append com.CreateParameter("@EmpNum", adChar, adParamInput, 10, UserNameReturn)
	
	if strWhichEntry = "Project" then
		com.CommandText = "spUpdateProjects"
		com.Parameters.Append com.CreateParameter("@strProjects", adVarChar, adParamInput, 5000, Trim(strString))
	elseif strWhichEntry = "Interest" then
		com.CommandText = "spUpdateInterests"
		com.Parameters.Append com.CreateParameter("@strInterests", adVarChar, adParamInput, 5000, Trim(strString))
	elseif strWhichEntry = "Details" then
		com.CommandText = "spUpdateDetails"
		com.Parameters.Append com.CreateParameter("@strPhoneNumber", adVarChar, adParamInput, 10, Trim(strPhoneNumber))
		com.Parameters.Append com.CreateParameter("@strRoomNumber", adVarChar, adParamInput, 6, Trim(strRoomNumber))
		com.Parameters.Append com.CreateParameter("@strAltContact", adVarChar, adParamInput, 10, Trim(strAltContact))
	end if
	
	com.Execute
	
end if

if com.State <> 0 then com.Close
set com = nothing
if con.State <> 0 then con.Close
set con = nothing

if Err.number = 0 then
	Response.Redirect "edithomepage.asp?Action=Save&Result=True"
else
	Response.Redirect "edithomepage.asp?Action=Save&Result=False&ErrNum=" & Err.number & "&ErrDesc" & Err.Description
end if
%>