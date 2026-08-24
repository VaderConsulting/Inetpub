<%
' Determines who is currently logged in

Dim WSHNetwork
Dim UserNameReturn
Dim PathNameReturn

' Retrieve who is currently logged in (ENum)
Set WSHNetwork = Server.CreateObject("WScript.Network")
UserNameReturn = WSHNetwork.UserName

' REMOVE AFTER TESTING '
'UserNameReturn = "c970"'
''''''''''''''''''''''''

' Retrieve directory
PathNameReturn = Request.ServerVariables("PATH_INFO")
PathNameReturn = mid(PathNameReturn,30,len(PathNameReturn)-41)

'Response.Write "Generated from whologgedin.asp: " & UserNameReturn & "<br>"
'Response.Write "Generated from whologgedin.asp: " & PathNameReturn

Set WSHNetwork = Nothing
%>