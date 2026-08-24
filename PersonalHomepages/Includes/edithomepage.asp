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
Dim AltContact
'**********************FUNCTIONS**************************

function displayField(strFieldName)
	' Displays a field
	On Error Resume Next
	
	if rst.state <> 0 then
		if NOT rst.eof then
			if not isnull(rst(strFieldName)) then
				strTemp = rst(strFieldName)
				strTemp = replace(strTemp,"<","&lt")
				strTemp = replace(strTemp,">","&gt")
			else
				strTemp = ""
			end if
			displayField = trim(strTemp)
		else
			displayField = "Unable to retrieve field"
		end if
	else
		displayField = "Unable to retrieve field"
	end if
	
	if Err.number <> 0 then
		displayField = "Err"
	end if
end function

sub retrieveDetails(strUserName)
	' Retrieves users details
	set rst = server.CreateObject("ADODB.Recordset")
	
	strSQL = "spRetrieveDetails '" & strUserName & "'"
	
	rst.Open strSQL,CONST_CONNECTION_STRING,adOpenStatic,adLockReadOnly
end sub

sub retrieveEmployeeList()
	' Retrieves users details
	if rst.state <> 0 then rst.close
	
	strSQL = "spRetrieveEmployeeList"
	
	rst.Open strSQL,CONST_CONNECTION_STRING,adOpenStatic,adLockReadOnly
end sub

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
			strStatus = "Unable to retrieve fields " & strUsername
		else
			strStatus = ""
			strProjects = rst("Projects")
			if isnull(strProjects) then strProjects = ""
			strInterests = rst("Interests")
			if isnull(strInterests) then strInterests = ""
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
retrieveDetails UserNameReturn

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

function saveDetails()
{
	// Check fields and opens save page
	strLocation = "save.asp?whichEntry=Details&PhoneNumber=" + escape(txtPhoneNumber.value) + "&RoomNumber=" + escape(txtRoomNumber.value) + "&AltContact=" + escape(cboAltContact.value)
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
		strProjects = removePipe(txtProjectName.value);
	}
	else
	{
		strProjects = spnComPro.innerText + "|" + removePipe(txtProjectName.value);
	}
	strLocation = "save.asp?whichEntry=Project&newEntry=" + escape(strProjects)
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
		strInterests = removePipe(txtInterestName.value) + "|" + removePipe(txtInterestURL.value);
	}
	else
	{
		strInterests = spnComInt.innerText + "|" + removePipe(txtInterestName.value) + "|" + removePipe(txtInterestURL.value);
	}
	strLocation = "save.asp?whichEntry=Interest&newEntry=" + escape(strInterests)
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

function showTitle()
{
	// Hide the Loading and show the title
	PleaseWait.className = "vis2"
	divTitle.className = "vis1"
}
</script>
<STYLE>
    .vis1 { visibility:visible }
    .vis2 { visibility:hidden }
</STYLE>


</head>
<BODY onLoad="javascript: showTitle()" bgColor="BAAF8D" Class="BodyText" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table width="100%" border="0" cellspacing="0" cellpadding="0" height="15" class="bodtext">
  <tr> 
    <td width="23%" height="49" rowspan="2" valign="top">
      <span id="PleaseWait">Loading Details...</span><div align="center" Class="vis2" id="divTitle"><img src="../Images/hpeditor.gif" width="130" height="23"></div>
    </td>
    <td width="1%" height="49" rowspan="2" valign="top"><img src="../Images/angle.gif" width="15" height="15"></td>
    <td width="76%" bgcolor="#FFFFFF" height="11"></td>
  </tr>
  <tr>
    <td width="76%" height="25"></td>
  </tr>
</table>
<%
if strStatus <> "" then
%>
<table width="100%" height="100%" cellpadding="0" cellspacing="0">
	<tr>
		<td class="WarningText" align="middle">
			The system was unable to retrieve your details from the database.<br>
			Please try again.<br>
			If this problem persists, please contact the System Administrator.<br>
			<%=strStatus%>
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
		<td>
			&nbsp
		</td>
	</tr>
	<tr>
    <td class="SectionHeaderText"> &nbsp;&nbsp;Personal Details</td>
	</tr>
	<tr>
		<td>&nbsp;
			
		</td>
	</tr>
	<tr> 
    <td> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="39">&nbsp;&nbsp;&nbsp;</td>
          <td width="500"> 
            <div align="left"></div>
            <table border="0" cellpadding="3" cellspacing="0" class="BodyText" align="left" style="BORDER-TOP: White 1px solid">
			  <tr class="ColumnHeaderText" style="BORDER-BOTTOM: White 1px solid"> 
                <td width="130" height="30" align="left" style="BORDER-BOTTOM: White 1px solid; BORDER-LEFT: White 1px solid"> 
                  Employee Number
                </td>
                <td height="30" align="left" style="BORDER-BOTTOM: White 1px solid; BORDER-LEFT: White 1px solid; BORDER-RIGHT: White 1px solid"> 
                  <%=displayField("EmpNum")%>
                </td>
			  </tr>
			  <tr class="ColumnHeaderText" style="BORDER-BOTTOM: White 1px solid"> 
                <td width="130" height="30" align="left" style="BORDER-BOTTOM: White 1px solid; BORDER-LEFT: White 1px solid"> 
                  Phone Number
                </td>
                <td height="30" align="left" style="BORDER-BOTTOM: White 1px solid; BORDER-LEFT: White 1px solid; BORDER-RIGHT: White 1px solid"> 
                  <input value="<%=displayField("PhoneNum")%>" id="txtPhoneNumber" style="WIDTH: 160px" class="textbox" name="Input" maxlength="20">
                </td>
			  </tr>
				<tr class="ColumnHeaderText" style="BORDER-BOTTOM: White 1px solid"> 
                <td width="130" height="30" align="left" style="BORDER-BOTTOM: White 1px solid; BORDER-LEFT: White 1px solid"> 
                  Room Number
                </td>
                <td height="30" align="left" style="BORDER-BOTTOM: White 1px solid; BORDER-LEFT: White 1px solid; BORDER-RIGHT: White 1px solid"> 
                  <input value="<%=displayField("RoomNum")%>" id="txtRoomNumber" style="WIDTH: 160px" class="textbox" name="Input" maxlength="6">
                </td>
			  </tr>
			  <tr class="ColumnHeaderText" style="BORDER-BOTTOM: White 1px solid"> 
                <td width="130" height="30" align="left" style="BORDER-BOTTOM: White 1px solid; BORDER-LEFT: White 1px solid"> 
                  Alternate Contact
                </td>
                <td height="30" align="left" style="BORDER-BOTTOM: White 1px solid; BORDER-left: White 1px solid; BORDER-RIGHT: White 1px solid"> 
                  <SELECT id="cboAltContact" style="WIDTH: 250px" class="textbox">
					<OPTION value="None">None</OPTION>
					<%
					AltContact = displayField("AltContactEmpNum")
					
					retrieveEmployeeList
					if rst.state <> 0 then
						do while NOT rst.eof
							if NOT isnull(AltContact) then
								if trim(rst("EmpNum")) = AltContact then
					%>
					<OPTION selected value="<%=rst("EmpNum")%>"><%=trim(rst("LastName")) & ", " & rst("PrefName") & " (" & trim(rst("EmpNum")) & ")"%></OPTION>
					<%
								else
					%>
					<OPTION value="<%=rst("EmpNum")%>"><%=trim(rst("LastName")) & ", " & rst("PrefName") & " (" & trim(rst("EmpNum")) & ")"%></OPTION>
					<%
								end if
							else
					%>
					<OPTION value="<%=rst("EmpNum")%>"><%=trim(rst("LastName")) & ", " & rst("PrefName") & " (" & trim(rst("EmpNum")) & ")"%></OPTION>
					<%
							end if
							rst.movenext
						loop
					end if
					%>
				  </SELECT>
                </td>
			  </tr>
			  <tr class="ColumnHeaderText"> 
                <td width="130" height="20" align="left" style=""> 
                  &nbsp
                </td>
                <td height="20" align="left"> 
                  &nbsp
                </td>
			  </tr>
			  <tr class="ColumnHeaderText"> 
                <td width="130" height="30" align="left" style=""> 
                  <input type="button" value="Save" id="btnSaveDetails" onClick="javascript:saveDetails()" name="button">
                </td>
                <td height="30" align="left"> 
                  &nbsp
                </td>
			  </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
	</tr>
	<tr>
		<td>&nbsp;
			
		</td>
	</tr>
	<tr>
	
	
	
	
	<tr>
		
    <td class="SectionHeaderText"> &nbsp;&nbsp;My Tasks </td>
	</tr>
	<tr>
		<td>&nbsp;
			
		</td>
	</tr>
	
  <tr> 
    <td> 
      <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="39">&nbsp;&nbsp;&nbsp;</td>
          <td width="500"> 
            <div align="left"></div>
            <table border="0" cellpadding="3" cellspacing="0" class="BodyText" align="center">
              <thead> 
              <tr class="ColumnHeaderText"> 
                <td width="47" height="30" align="middle" style="BORDER-BOTTOM: White 1px solid"> 
                  Delete </td>
                <td height="30" style="BORDER-BOTTOM: White 1px solid" width="573"> 
                  Name </td>
              </tr>
              </thead> <tbody> 
              <%
					fillArray "Projects"
					strTemp =""
					if ubound(arystrSplit) <> -1 then
						for i = 0 to ubound(arystrSplit) step 1
							if strTemp = "" then
								strTemp = arystrSplit(i)
							else
								strTemp = strTemp & "|" & arystrSplit(i)
							end if
						%>
              <tr> 
                <td height="30" width="47" class="SmallText" align="middle" style="BORDER-BOTTOM: White 1px solid; BORDER-RIGHT: White 1px solid; BORDER-LEFT: White 1px solid"> 
                  <a href="javascript:deleteProject(<%=i%>)">Delete</a> 
                </td>
                <td height="30" width="573" style="BORDER-BOTTOM: White 1px solid; BORDER-RIGHT: White 1px solid"> 
                  <span style="OVERFLOW-X: auto; WIDTH: 570px; HEIGHT: 30px"><%=arystrSplit(i)%></span> </td>
              </tr>
              <%
						next
					else
						%>
              <tr> 
                <td colspan="3" height="30" class="SmallText" style="BORDER-BOTTOM: White 1px solid; BORDER-RIGHT: White 1px solid; BORDER-LEFT: White 1px solid"> 
                  &nbsp&nbsp&nbspYou have no Tasks recorded</td>
              </tr>
              <%
					end if
					%>
              </tbody> <tfoot> 
              <tr> 
                <td colspan="3">&nbsp; </td>
              </tr>
              <tr class="ColumnHeaderText"> 
                <td colspan="3"> Add new project<span id="spnComPro" class="InvisibleSpan"><%=strTemp%></span> </td>
              </tr>
              <tr class="ColumnHeaderText"> 
                <td width="47">&nbsp; </td>
                <td width="573"> Name </td>
              </tr>
              <tr class="ColumnHeaderText"> 
                <td width="47"> 
                  <input type="button" value="Add" id="btnNewProject" onClick="javascript:saveProject()" name="button">
                </td>
                <td width="573"> 
                  <input id="txtProjectName" style="WIDTH: 190px" class="textbox" name="Input">
                </td>
              </tr>
              </tfoot> 
            </table>
          </td>
        </tr>
      </table>
    </td>
	</tr>
	<tr>
		<td>&nbsp;
			
		</td>
	</tr>
	<tr>
		
    <td class="SectionHeaderText"> &nbsp;&nbsp;Interests </td>
	</tr>
	<tr>
		<td>&nbsp;
			
		</td>
	</tr>
	<tr>
		<td>
			      <table border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="39">&nbsp;&nbsp;&nbsp;</td>
          <td width="500"> 
            <div align="left"></div>
<table border="0" cellpadding="3" cellspacing="0" class="BodyText">
				<thead>
					<tr class="ColumnHeaderText">
						<td width="60" align="middle" height="30" style="BORDER-BOTTOM: White 1px solid">
							Delete
						</td>
						<td height="30" style="BORDER-BOTTOM: White 1px solid">
							Name
						</td>
						<td height="30" style="BORDER-BOTTOM: White 1px solid">
							URL
						</td>
					</tr>
				</thead>
				<tbody>
					<%
					fillArray "Interests"
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
							<td width="60" class="SmallText" align="middle" height="30" style="BORDER-BOTTOM: White 1px solid; BORDER-LEFT: White 1px solid; BORDER-RIGHT: White 1px solid">
								<a href="javascript:deleteInterest(<%=i%>)">Delete</a>
							</td>
							<td height="30" style="BORDER-BOTTOM: White 1px solid; BORDER-RIGHT: White 1px solid">
								<span style="OVERFLOW-X: auto; WIDTH: 200px; HEIGHT: 30px"><%=arystrSplit(i)%></span>
							</td>
							<td height="30" style="BORDER-BOTTOM: White 1px solid; BORDER-RIGHT: White 1px solid">
								<span style="OVERFLOW-X: auto; WIDTH: 300px; HEIGHT: 30px"><%=arystrSplit(i+1)%></span>
							</td>
						</tr>
						<%
						next
					else
						%>
						<tr>
							<td colspan="4" height="30" class="SmallText" style="BORDER-BOTTOM: White 1px solid; BORDER-RIGHT: White 1px solid; BORDER-LEFT: White 1px solid">
								&nbsp&nbsp&nbspYou have no Interests recorded.
							</td>
						</tr>
						<%
					end if
					%>
				</tbody>
				<tfoot>
					<tr>
						<td colspan="4">&nbsp;
							
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td colspan="4">
							Add new interest<span id="spnComInt" class="InvisibleSpan"><%=strTemp%></span>
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td>&nbsp;
							
						</td>
                        <td>Name
							
						</td>
						<td nowrap>
							URL <span style="bodtext">(include http:// if external site)</span>
						</td>
					</tr>
					<tr class="ColumnHeaderText">
						<td width="60">
							<INPUT type="button" value="Add" id="btnNewInterest" onClick="javascript:saveInterests()">
						</td>
                        <td>
							<INPUT id="txtInterestName" style="WIDTH: 190px" class="textbox">
						</td>
						<td>
							<INPUT id="txtInterestURL" style="WIDTH: 190px" class="textbox">
						</td>
					</tr>
				</tfoot>
			</table>
		          </td>
        </tr>
      </table>
		</td>
	</tr>
	
	<tr>
		<td height="100%">&nbsp;
			
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
	if rst.state <> 0 then rst.close
	set rst = nothing
%>