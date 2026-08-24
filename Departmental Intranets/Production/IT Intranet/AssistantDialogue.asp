<%@Language = VBScript%>
<%
  Option Explicit
  
	Dim ServiceAccount
	ServiceAccount = Application("ADServiceAccount")
	Dim ServicePassword
	ServicePassword = Application("ADServicePassword")

  ' If the last time this page was executed was over an hour ago (or never) create the assistant HTML list and cache it in the Application("AssistantListHTML") field
  Dim dteLastAssistantRefresh
  dteLastAssistantRefresh = Application("LastAssistantRefreshDateTime")
  If IsEmpty(dteLastAssistantRefresh) Then dteLastAssistantRefresh = DateSerial(1900,1,1)
  If IsEmpty(Application("AssistantListHTML")) Then dteLastAssistantRefresh = DateSerial(1900,1,1)
  If (DateAdd("h",-1,Now) > dteLastAssistantRefresh) Then 'Or True Then
    Application.Lock 
    Application("LastAssistantRefreshDateTime")= Now
    Response.Write "<!-- Started Creating Assistant List " + CStr(Now) + " -->"
    Application("AssistantListHTML") = GenerateAssistantOptionList
    Response.Write "<!-- Finished Creating Assistant List " + CStr(Now) + " -->"
    Application.Unlock
  Else
    Response.Write "<!-- Using cached results from " + CStr(dteLastAssistantRefresh) + " -->"
  End If

Function GenerateAssistantOptionList
	'define the recordset objects
	Dim adoConn
	Dim adoCmd
	Dim adoRS
  Dim arrResults()
  ReDim arrResults(2,0)
  ' 0 is the sorting key
  ' 1 is the display text
  ' 2 is the DN
  
  Dim intCount,intCount2,intCount3
  Dim strMisc
  Dim strOptionHTML
  Dim lngUAC

	'create the connection object
	Set adoConn = CreateObject("ADODB.Connection")
	
	'Attempt to locate the user account in AD
	adoConn.Provider = "ADsDSOObject"
	adoConn.Properties("User ID") = ServiceAccount
	adoConn.Properties("Password") = ServicePassword
	adoConn.Properties("ADSI Flag") = 1
	adoConn.Properties("Encrypt Password") = True
	adoConn.Open "ADSI"

	Set adoCmd = CreateObject("ADODB.Command")
	Set adoCmd.ActiveConnection = adoConn
	adoCmd.Properties("Cache results") = True
	adoCmd.Properties("Page Size") = 1000
	adoCmd.CommandText = "<LDAP://OU=User Accounts,DC=Oxiana,DC=com,DC=au>;(&(objectCategory=user)(objectClass=user));userAccountControl,canonicalName,distinguishedname,givenname,sn,title;subtree"
	Set adoRS = adoCmd.Execute
  intCount = 0
	While Not adoRS.EOF
	  lngUAC = CLng(adoRS("userAccountControl"))
	  ' Sanity Check
	  ' 1. Surname must be longer than 0 chars
	  ' 2. Given Name must be longer than 0 chars
	  ' 3. Account must not be disabled
		If (Len(adoRS.Fields("SN")) > 0) And (Len(adoRS.Fields("GivenName")) > 0) And ((lngUAC And 2) = 0) Then
      If LBound(arrResults,2) < intCount Then
        ' Increment the array
        ReDim Preserve arrResults(2,intCount)
      End If
      ' Display Text
      arrResults(1,intCount) = CStr(adoRS.Fields("sn") & ", " & adoRS.Fields("givenname"))
      If Not IsNull(adoRS("title")) Then arrResults(1,intCount) = arrResults(1,intCount) + " (" + CStr(adoRS("title")) + ")"
      ' Sorting Text
      arrResults(0,intCount) = UCase(arrResults(1,intCount))
      ' DN
      arrResults(2,intCount) = CStr(adoRS("distinguishedname"))

      intCount = intCount + 1
    End If
	  adoRS.MoveNext
	Wend
	adoRS.Close

  ' Bubble Sort it....Sure it takes a while but it's quick enough at the moment.
  For intCount = 0 To UBound(arrResults,2) - 1
    For intCount2 = intCount To UBound(arrResults,2)
      If arrResults(0,intCount) > arrResults(0,intCount2) Then
        For intCount3 = 0 To 2
          strMisc = arrResults(intCount3,intCount)
          arrResults(intCount3,intCount) = arrResults(intCount3,intCount2)
          arrResults(intCount3,intCount2) = strMisc
        Next
      End If
    Next
  Next

  strOptionHTML = ""
	For intCount = LBound(arrResults,2) To UBound(arrResults,2)
    strOptionHTML = strOptionHTML + "<OPTION value=""" + arrResults(2,intCount) + """>" + Server.HTMLEncode(arrResults(1,intCount)) + "</OPTION>"
  Next

  GenerateAssistantOptionList = strOptionHTML
End Function
%>
<HTML>
<HEAD>
	<TITLE>Select an assistant...</TITLE>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
	<SCRIPT language=JavaScript>
	<!--
		function CheckSelected()
		{
			var Counter;
			for (Counter=0;Counter<Assistant.options.length;Counter++)
			{
				if (Assistant.options[Counter].value == window.dialogArguments)
				{
					Assistant.options[Counter].selected = true;
					return true;
				}
			}
		}
		function SelectAssistant()
		{
			if (Assistant.selectedIndex < 0)
			{
				alert('Please select an assistant from the list first');
			}
			else
			{
				window.returnValue = Assistant.options(Assistant.selectedIndex).value + '|' + Assistant.options(Assistant.selectedIndex).text;
				window.close();
			}
		}
	//-->
	</SCRIPT>
</HEAD>

<BODY onload="CheckSelected()">
	<TABLE cellpadding=0 cellspacing=0 width=400>
		<TR><TD>&nbsp;</TD></TR>
		<TR>
			<TD align=center class=SubHeading>
				Please select an assistant from the list:
			</TD>
		</TR>
		<TR>
			<TD align=center>
				<SELECT style="width:380px;" size=10 id=Assistant><%= Application("AssistantListHTML") %></SELECT>
			</TD>
		</TR>
		<TR><TD>&nbsp;</TD></TR>
		<TR>
			<TD align=center>
				<INPUT type=button value="OK" style="WIDTH:75px" onclick="SelectAssistant()" id=button1 name=button1>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=button value="Cancel" style="WIDTH:75px" onclick="window.close()" id=button2 name=button2>
			</TD>
		</TR>
	</TABLE>
	
</BODY>
</HTML>