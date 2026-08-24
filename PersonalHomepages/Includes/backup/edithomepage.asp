<%@ Language=VBScript %>
<%'ABOVE LINE CAN BE REMOVED IF A SESSION IS REQUIRED
Option Explicit 
Response.Expires = true
'***************************************************************************
' Name:		edithomepage.asp
' Purpose:	Allow user to edit Projects and Interests fields
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
Dim rst
Dim strStatus
Dim strProjects ' Name|Section
Dim strInterests ' Name|Description|URL
Dim arystrSplit
Dim blnResult
Dim strAction
Dim strErrNum
Dim strErrDesc
Dim strTemp
'**********************FUNCTIONS**************************

sub fillArray(strType)
	' Uses the split function to fill an array
	if lcase(strType) = "projects" then
		arystrSplit = split(strProjects,"|")
	elseif lcase(strType) = "interests" then
		arystrSplit = split(strInterests,"|")
	end if
end sub

sub retrieveFields(strUserName)
	' Retrieves fields Projects and Interests
	' Fields: Projects, Interests
	
	set rst = server.CreateObject("ADODB.Recordset")
	
	strSQL = "spRetrieveProjectsInterests '" & strUserName & "'"
	
	rst.Open strSQL,CONST_CONNECTION_STRING,adOpenStatic,adLockReadOnly
	
	if rst.State <> 0 then
		if rst.EOF then
			strStatus = "Unable to retrieve fields"
		else
			strStatus = ""
			strProjects = rst("Projects")
			strInterests = rst("Interests")
		end if
	else
		strStatus = "Unable to retrieve fields"
	end if
	
	if rst.State <> 0 then rst.Close
	set rst = nothing
end sub

function fixQuotes(strString)
	' Fixes any quote problems
	Dim strFixed
	
	if not isnull(strString) then
		strFixed = replace(strString,"""","""""")
		fixQuotes = replace(strFixed,"'","''")
	else
		fixQuotes = strString
	end if
end function

'*************************MAIN****************************
' Retrieve any query string variables
strAction = Request.QueryString("Action")
blnResult = Request.QueryString("Result")

if blnResult <> true then
	strErrNum = Request.QueryString("ErrNum")
	strErrDesc = Request.QueryString("ErrDesc")
end if

' Filter contractors if search criteria supplied
retrieveFields UserNameReturn

%>
<html>
<head>
<title>Edit Personal Homepage</title>
<link rel="stylesheet" href="style.css">
<style type="text/css"></style>

<script language="javascript">

function removePipe(strString)
{
	var strNewString = "";
	
	// Loop through string and replace | with nothing
	
	
	
	return strString;
}
		
function deleteProject(intNumber)
{
	// Deletes Project
	strLocation = "deleteprojectinterest.asp?whichEntry=Project&whichNumber=" + intNumber + "&allString=" + spnComPro.innerText;
	location.replace(strLocation);
}

function deleteInterest(intNumber)
{
	// Deletes Interest
	strLocation = "deleteprojectinterest.asp?whichEntry=Interest&whichNumber=" + intNumber + "&allString=" + spnComInt.innerText;
	location.replace(strLocation);
}

function saveProject()
{
	// Check fields and opens save page
	if (trim_string(txtProjectName.value) == '')
	{
		alert("You must enter a project name.");
		txtProjectName.value = "";
		txtProjectName.select();
		return false;
	}
	if (spnComPro.innerText == '')
	{
		strProjects = removePipe(txtProjectName.value) + "|" + removePipe(txtProjectSection.value);
	}
	else
	{
		strProjects = spnComPro.innerText + "|" + removePipe(txtProjectName.value) + "|" + removePipe(txtProjectSection.value);
	}
	strLocation = "saveprojectinterest.asp?whichEntry=Project&newEntry=" + escape(strProjects)
	location.replace(strLocation);
}

function saveInterests()
{
	// Check fields and opens save page
	if (trim_string(txtInterestName.value) == '')
	{
		alert("You must enter an interest name.");
		txtInterestName.value = "";
		txtInterestName.select();
		return false;
	}
	if (spnComInt.innerText == '')
	{
		strInterests = removePipe(txtInterestName.value) + "|" + removePipe(txtInterestDescription.value) + "|" + removePipe(txtInterestURL.value);
	}
	else
	{
		strInterests = spnComInt.innerText + "|" + removePipe(txtInterestName.value) + "|" + removePipe(txtInterestDescription.value) + "|" + removePipe(txtInterestURL.value);
	}
	strLocation = "saveprojectinterest.asp?whichEntry=Interest&newEntry=" + escape(strInterests)
	location.replace(strLocation);
}

// Trims a string
function trim_string(strValue) 
{
	var ichar, icount;
	//var strValue = this;
	ichar = strValue.length - 1;
	icount = -1;
	while (strValue.charAt(ichar)==' ' && ichar > icount)
		--ichar;
	if (ichar!=(strValue.length-1))
		strValue = strValue.slice(0,ichar+1);
	ichar = 0;
	icount = strValue.length - 1;
	while (strValue.charAt(ichar)==' ' && ichar < icount)
		++ichar;
	if (ichar!=0)
		strValue = strValue.slice(ichar,strValue.length);
	return strValue;
}
 
</script>

</head>
<BODY bgColor="silver" Class="BodyText">

<%
if strStatus <> "" then
%>
<table width="100%" height="100%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="WarningText" align="middle">
			The system was unable to retrieve your details from the database.<br>
			Please try again.<br>
			If this problem persists, please contact the System Administrator.
		</td>
	</tr>
</table>
<%
else

	if strAction = "Save" AND blnResult = "True" then
%>

The record was saved successfully.<br><br>

	<%
	elseif strAction = "Save" AND blnResult = "False" then
	%>

There was an error saving the record.<br>
Please try again.<br>
If this problem persists, please contact the System Administrator.<br><br>

Error Number: <%=strErrNum%><br>
Error Description: <%=strErrDesc%><br><br>

	<%
	end if
	if strAction = "Delete" AND blnResult = "True" then
%>

The record was deleted successfully.<br><br>

	<%
	elseif strAction = "Delete" AND blnResult = "False" then
	%>

There was an error deleting the record.<br>
Please try again.<br>
If this problem persists, please contact the System Administrator.<br><br>

Error Number: <%=strErrNum%><br>
Error Description: <%=strErrDesc%><br><br>

	<%
	end if
	%>

<table border="0" width="100%" height="100%" cellpadding="0" cellspacing="0" class="BodyText">
	<tr>
		<td class="SectionHeaderText">
			Projects
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;
		</td>
	</tr>
	<tr>
		<td>
			<table border="0" cellpadding="3" cellspacing="0" class="BodyText">
				<thead>
					<tr class="ColumnHeaderText">
						<td width="60" height="30" align="middle" style="BORDER-BOTTOM: black 1px solid">
							Delete
						</td>
						<td height="30" style="BORDER-BOTTOM: black 1px solid">
							Name
						</td>
						<td height="30" style="BORDER-BOTTOM: black 1px solid">
							Section
						</td>
					</tr>
				</thead>
				<tbody>
					<%
					fillArray "Projects"
					strTemp =""
					if ubound(arystrSplit) <> -1 then
						for i = 0 to ubound(arystrSplit) step 2
							if strTemp = "" then
								strTemp = arystrSplit(i) & "|" & arystrSplit(i+1)
							else
								strTemp = strTemp & "|" & arystrSplit(i) & "|" & arystrSplit(i+1)
							end if
						%>
						<tr>
							<td height="30" width="60" class="SmallText" align="middle" style="BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid; BORDER-LEFT: black 1px solid">
								<a href="javascript:deleteProject(<%=i%>)">Delete</a>
							</td>
							<td height="30" width="200" style="BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid">
								<span style="OVERFLOW-X: auto; WIDTH: 200px; HEIGHT: 30px"><%=arystrSplit(i)%></span>
							</td>
							<td height="30" style="BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid">
								<span style="OVERFLOW-X: auto; WIDTH: 200px; HEIGHT: 30px"><%=arystrSplit(i+1)%></span>
							</td>
						</tr>
						<%
						next
					else
						%>
						<tr>
							<td colspan="3" height="30" class="SmallText" style="BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid; BORDER-LEFT: black 1px solid">
								&nbsp&nbsp&nbspYou have no Projects
							</td>
						</tr>
						<%
					end if
					%>
				</tbody>
				<tfoot>
					<tr>
						<td colspan="3">
							&nbsp;
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td colspan="3">
							Add new project<span id="spnComPro" class="InvisibleSpan"><%=strTemp%></span>
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td>
							&nbsp
						</td>
						<td>
							Name
						</td>
						<td>
							Section
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td width="60">
							<INPUT type="button" value="Add" id="btnNewProject" onClick="javascript:saveProject()">
						</td>
						<td>
							<INPUT id="txtProjectName" maxlength="50" style="WIDTH: 190px">
						</td>
						<td>
							<INPUT id="txtProjectSection" maxlength="50" style="WIDTH: 190px">
						</td>
					</tr>
				</tfoot>
			</table>
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;
		</td>
	</tr>
	<tr>
		<td class="SectionHeaderText">
			Interests
		</td>
	</tr>
	<tr>
		<td>
			&nbsp;
		</td>
	</tr>
	<tr>
		<td>
			<table border="0" cellpadding="3" cellspacing="0" class="BodyText">
				<thead>
					<tr class="ColumnHeaderText">
						<td width="60" align="middle" height="30" style="BORDER-BOTTOM: black 1px solid">
							Delete
						</td>
						<td height="30" style="BORDER-BOTTOM: black 1px solid">
							Name
						</td>
						<td height="30" style="BORDER-BOTTOM: black 1px solid">
							Description
						</td>
						<td height="30" style="BORDER-BOTTOM: black 1px solid">
							URL
						</td>
					</tr>
				</thead>
				<tbody>
					<%
					fillArray "Interests"
					strTemp =""
					if ubound(arystrSplit) <> -1 then
						for i = 0 to ubound(arystrSplit) step 3
							if strTemp = "" then
								strTemp = arystrSplit(i) & "|" & arystrSplit(i+1) & "|" & arystrSplit(i+2)
							else
								strTemp = strTemp & "|" & arystrSplit(i) & "|" & arystrSplit(i+1) & "|" & arystrSplit(i+2)
							end if
						%>
						<tr>
							<td width="60" class="SmallText" align="middle" height="30" style="BORDER-BOTTOM: black 1px solid; BORDER-LEFT: black 1px solid; BORDER-RIGHT: black 1px solid">
								<a href="javascript:deleteInterest(<%=i%>)">Delete</a>
							</td>
							<td height="30" style="BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid">
								<span style="OVERFLOW-X: auto; WIDTH: 200px; HEIGHT: 30px"><%=arystrSplit(i)%></span>
							</td>
							<td height="30" style="BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid">
								<span style="OVERFLOW-X: auto; WIDTH: 200px; HEIGHT: 30px"><%=arystrSplit(i+1)%></span>
							</td>
							<td height="30" style="BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid">
								<span style="OVERFLOW-X: auto; WIDTH: 200px; HEIGHT: 30px"><%=arystrSplit(i+2)%></span>
							</td>
						</tr>
						<%
						next
					else
						%>
						<tr>
							<td colspan="4" height="30" class="SmallText" style="BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid; BORDER-LEFT: black 1px solid">
								&nbsp&nbsp&nbspYou have no Interests
							</td>
						</tr>
						<%
					end if
					%>
				</tbody>
				<tfoot>
					<tr>
						<td colspan="4">
							&nbsp;
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td colspan="4">
							Add new interest<span id="spnComInt" class="InvisibleSpan"><%=strTemp%></span>
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td>
							&nbsp
						</td>
						<td>
							Name
						</td>
						<td>
							Description
						</td>
						<td>
							URL
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td width="60">
							<INPUT type="button" value="Add" id="btnNewInterest" onClick="javascript:saveInterests()">
						</td>
						<td>
							<INPUT id="txtInterestName" maxlength="50" style="WIDTH: 190px">
						</td>
						<td>
							<INPUT id="txtInterestDescription" maxlength="50" style="WIDTH: 190px">
						</td>
						<td>
							<INPUT id="txtInterestURL" maxlength="50" style="WIDTH: 190px">
						</td>
					</tr>
				</tfoot>
			</table>
		</td>
	</tr>
	
	<tr>
		<td height="100%">
			&nbsp;
		</td>
	</tr>
</table>

<%
end if
%>

</BODY>
</html>
<%
' Clean up

%>