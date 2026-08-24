<%
' Retrieves the Projects and Interests fields from the DOS database

Dim WSHNetwork
Dim UserNameReturn
Dim PathNameReturn

' Retrieve who is currently logged in (ENum)
Set WSHNetwork = Server.CreateObject("WScript.Network")
UserNameReturn = WSHNetwork.UserName

' Retrieve directory
PathNameReturn = Request.ServerVariables("PATH_INFO")
PathNameReturn = mid(PathNameReturn,30,len(PathNameReturn)-41)

Response.Write UserNameReturn & "<br>"
Response.Write PathNameReturn

Set WSHNetwork = Nothing
%>