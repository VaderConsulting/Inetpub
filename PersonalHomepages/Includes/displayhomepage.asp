<%'ABOVE LINE CAN BE REMOVED IF A SESSION IS REQUIRED
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
' NOTE The term PROJECTS has been redefined to mean 'My Tasks'.
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

sub fillArray(strType)
	' Uses the split function to fill an array
	if lcase(strType) = "projects" then
		arystrSplit = split(strProjects,"|")
	elseif lcase(strType) = "interests" then
		arystrSplit = split(strInterests,"|")
	end if
end sub

sub retrieveDetails(strUserName)
	' Retrieves users details
	set rst = server.CreateObject("ADODB.Recordset")
	
	strSQL = "spRetrieveDetails '" & strUserName & "'"
	
	rst.Open strSQL,CONST_CONNECTION_STRING,adOpenStatic,adLockReadOnly
end sub

sub retrieveIPFields(strUserName)
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

function ensureHTTP(strString)
	' Ensures there is http:// at the start of a URL
	
	if lcase(left(strString,7)) <> "http://" then
		ensureHTTP = "http://" & strString
	else
		ensureHTTP = strString
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

' Retrieve fields
if UserNameReturn = PathNameReturn then
	retrieveIPFields UserNameReturn
else
	retrieveIPFields PathNameReturn
end if

' Retrieve details
if UserNameReturn = PathNameReturn then
	retrieveDetails UserNameReturn
else
	retrieveDetails PathNameReturn
end if

if strStatus <> "" then
%>
<table width="100%" height="100" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="middle"> The system was unable to retrieve your details from the 
      database.<br>
      Please try again.<br>
      If this problem persists, please contact the System Administrator. </td>
  </tr>
</table>
<%
else

	if strAction = "Save" AND blnResult = "True" then
%>
The record was saved successfully.<br>
<br>
<%
	elseif strAction = "Save" AND blnResult = "False" then
	%>
There was an error saving the record.<br>
Please try again.<br>
If this problem persists, please contact the System Administrator.<br>
<br>
Error Number: <%=strErrNum%><br>
Error Description: <%=strErrDesc%><br>
<br>
<%
	end if
	if strAction = "Delete" AND blnResult = "True" then
%>
The record was deleted successfully.<br>
<br>
<%
	elseif strAction = "Delete" AND blnResult = "False" then
	%>
There was an error deleting the record.<br>
Please try again.<br>
If this problem persists, please contact the System Administrator.<br>
<br>
Error Number: <%=strErrNum%><br>
Error Description: <%=strErrDesc%><br>
<br>
<%
	end if
	%>
<table width="80%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td width="2%">&nbsp;</td>
    <td width="50%"> 
      <table width="700" border="0" cellspacing="0" class="blockbg" cellpadding="0" bgcolor="#F0F0F0" style="BORDER-TOP: black 1px solid; BORDER-BOTTOM: black 1px solid; BORDER-RIGHT: black 1px solid; BORDER-LEFT: black 1px solid">
        <tr> 
          <td width="48%" valign="top"> 
            <div align="left"></div>
            <br>
            <table width="93%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr> 
                <td height="107" width="42%" align="center" valign="top"> 
                  <div align="center"><img src="images/id.gif" width="125" height="155" border="0"></div>
                </td>
                <td height="107" width="58%" valign="top"> 
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr class="bodtext"> 
                      <td width="250"><b><font size="4"><%=displayField("PrefName") & " " & displayField("LastName")%></font></b></td>
                    </tr>
                    <tr class="bodtext"> 
                      <td width="250"><a href="<%="../default.asp?Position=" & Server.URLEncode(displayField("Position"))%>&cmdGo=Go"><%=displayField("Position") & "</a> " & displayField("Acronym")%></td>
                    </tr>
                    <tr class="bodtext"> 
                      <td width="250"><%=displayField("PhoneNum")%></td>
                    </tr>
                    <tr class="bodtext"> 
                      <td width="250"> 
                        <div align="left"><%=displayField("EmpNum")%></div>
                      </td>
                    </tr>
                    <tr class="bodtext"> 
                      <td width="250"><a href="<%="../default.asp?Location=" & Server.URLEncode(displayField("SiteName"))%>&cmdGo=Go"><%=displayField("SiteName")%></a>&nbsp;<%=displayField("RoomNum")%></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                    </tr>
                    <tr class="bodtext"> 
                      <%
                      LeaveStart = displayField("LeaveStart") & ""
                      LeaveEnd = displayField("LeaveEnd") & ""
                      If LeaveStart <> "" and LeaveEnd <> "" Then%>
                          <%If CDate(LeaveStart) < now and CDate(LeaveEnd) > now Then%>
                              <td width="250"><font size="1" color="#ff0000">On Leave Until <%=LeaveEnd%></td>
                          <%Else%>
                              <td width="250">&nbsp;</td>
                          <%End If%>
                      <%Else%>
                          <td width="250">&nbsp;</td>
                      <%End If%>
                    </tr>
                  </table>
                </td>
              </tr>
            </table>
            <br>
            
            <br>
          </td>
          <td width="50%" valign="top"><br><br>
            <table width="50%" border="0" cellspacing="0" cellpadding="0">
              <tr class="bodtext"> 
                <td>&nbsp;Alternate Contact</td>
              </tr>
            </table>
            <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr class="bodtext"> 
                <td><b><a href="<%="../create.asp?ENumber=" & displayField("AltContactEmpNum")%>"><%=displayField("AltContactPrefName") & " " & displayField("AltContactLastName")%></a></b>&nbsp</td>
              </tr>
            </table>
            <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr class="bodtext"> 
                <td><b><%=displayField("AltContactPosition") & " " & displayField("AltContactAcronym")%></b></td>
              </tr>
            </table>
            <table width="50%" border="0" cellspacing="0" cellpadding="0">
              <tr class="bodtext"> 
                <td>&nbsp;Branch</td>
              </tr>
            </table>
            <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr class="bodtext"> 
                <td><b><a href="<%="../default.asp?section=" & Server.URLEncode(displayField("SectionName")) & "&cmdGo=Go"%>"><%=displayField("SectionName")%></a></b>&nbsp</td>
              </tr>
            </table>
            <table width="50%" border="0" cellspacing="0" cellpadding="0">
              <tr class="bodtext"> 
                <td>&nbsp;Directorate</td>
              </tr>
            </table>
            <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr class="bodtext"> 
                <td><b>&nbsp</b></td>
              </tr>
            </table>
            <table width="50%" border="0" cellspacing="0" cellpadding="0">
              <tr class="bodtext"> 
                <td>&nbsp;Manager</td>
              </tr>
            </table>
            <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr class="bodtext"> 
                <td><b><a href="../create.asp?ENumber=<%=displayField("ManagerEmpNum")%>"><%=displayField("ManagerPrefName") & " " & displayField("ManagerLastName")%></a></b>&nbsp</td>
              </tr>
            </table>
            <table width="80%" border="0" cellspacing="0" cellpadding="0" align="center">
              <tr class="bodtext"> 
                <td><b><%=displayField("ManagerPosition") & " " & displayField("ManagerAcronym")%></b></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
			<td colspan="2">
			<table width="25%" border="0" cellspacing="0" cellpadding="0">
              <tr class="bodtext"> 
                <td>
                  <div>&nbsp;&nbsp;&nbsp;My Tasks</div>
                </td>
              </tr>
            </table>
            <table border="0" cellpadding="3" cellspacing="0" class="bodtext" align="center" width="95%">
              <tbody> 
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
                <td width="356"> <span style="OVERFLOW-X: auto"><b><%=arystrSplit(i)%></b></span> </td>
              </tr>
              <%
						next
					else
						%>
              <tr> 
                <td colspan="3" width="356"> <b>There are no Projects recorded</b> </td>
              </tr>
              <%
					end if
					%>
              </tbody> 
            </table>
            <br>
            <table width="25%" border="0" cellspacing="0" cellpadding="0">
              <tr class="bodtext"> 
                <td>
                  <div>&nbsp;&nbsp;&nbsp;My Interests</div>
                </td>
              </tr>
            </table>
            <table border="0" cellpadding="3" cellspacing="0" class="bodtext" align="center" width="95%">
              <tbody> 
              <%
					fillArray "Interests"
					strTemp =""
					on error resume next
					if ubound(arystrSplit) <> -1 then
						for i = 0 to ubound(arystrSplit) step 2
							if strTemp = "" then
								strTemp = arystrSplit(i) & "|" & arystrSplit(i+1)
							else
								strTemp = strTemp & "|" & arystrSplit(i) & "|" & arystrSplit(i+1)
							end if
						%>
              <tr> 
                <td> 
                  <%
									if arystrSplit(i+1) = "" or isnull(arystrSplit(i+1)) then
								%>
                  <span style="OVERFLOW-X: auto"><b><%=arystrSplit(i)%></b></span> 
                  <%
									else
								%>
                  <a target="_blank" href="<%=arystrSplit(i+1)%>"><span style="OVERFLOW-X: auto; Cursor: Hand"><b><%=arystrSplit(i)%></b></span></a> 
                  <%
									end if
								%>
                </td>
              </tr>
              <%
						next
					else
						%>
              <tr> 
                <td colspan="4" height="32"> <b>There are no Interests recorded</b> </td>
              </tr>
              <%
					end if
					%>
              </tbody> 
            </table>
			</td>
		</tr>
      </table>
      <form action="../default.asp" name="frmSearchReturn" id="frmSearchReturn">
      <table align="right">
      <tr>
      <td><input type="Submit" value="Search"</td>
      </tr>
      </table>
      </form>
    </td>
  </tr>
</table>
<%
end if

' Clean up
if rst.state <> 0 then rst.close
set rst = nothing
%>
