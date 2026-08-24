<%
' Determines who is currently logged in

'Dim WSHNetwork
Dim UserNameReturn
Dim PathNameReturnFull
Dim PathNameReturn
Dim Usr

' Retrieve who is currently logged in (ENum)
'Set WSHNetwork = Server.CreateObject("WScript.Network")
'UserNameReturn = WSHNetwork.UserName
'Get Username
Usr = UCase(Request.ServerVariables("AUTH_USER"))
i = InStr(Usr,"\")
Session("Username") = Mid(Usr,i+1,Len(Usr)-i)
UserNameReturn = Session("Username")

' Retrieve directory
PathNameReturnFull = Request.ServerVariables("PATH_INFO")
j = 0
for i = len(PathNameReturnFull) to 0 step -1
	if mid(PathNameReturnFull,i,1) = "/" then
		if j = 0 then
			j = i
		else
			PathNameReturn = mid(PathNameReturnFull,i+1,j-i-1)
			exit for
		end if
	end if
next

'Response.Write "Generated from whologgedin.asp: " & UserNameReturn & "<br>"
'Response.Write "Generated from whologgedin.asp: " & PathNameReturnFull & "<br>"
'Response.Write "Generated from whologgedin.asp: " & PathNameReturn

'Set WSHNetwork = Nothing
%>
