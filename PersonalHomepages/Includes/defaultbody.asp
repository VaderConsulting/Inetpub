<%@ Language=VBScript %>
<%
option explicit
Response.Buffer = True
'On Error Resume Next
%>
<!--#include file="constants.asp"-->
<!--#include file="whologgedin.asp"-->
<%
' Own homepage
if lcase(Session("username")) = lcase(PathNameReturn) then
	Response.Write "<frameset framespacing=""1"" border=""0"" frameborder=""NO"" rows=""90%,10%"" bordercolor=""#FFFFFF"">"
	Response.Write "	<frame name=""frmMain"" src=""homepage.asp""     scrolling=""yes"" frameborder=""1"" Title=""Personal Web Site"">"
	Response.Write "	<frame name=""frmEdit"" src=""../../includes/edithomepage.asp"" scrolling=""yes"" frameborder=""4"" Title=""Edit Personal Web Site"">"
	Response.Write "</frameset>"
	Response.Write "<noframes>"
	Response.Write "<body>"
	Response.Write "<p>This page uses frames, but your browser doesn't support them.</p>"
	Response.Write "</body>"
	Response.Write "</noframes>"
else
' Others homepage
	Response.Write "<frameset framespacing=""1"" border=""0"" frameborder=""NO"" rows=""100%"" bordercolor=""#FFFFFF"">"
	Response.Write "	<frame name=""frmMain"" src=""homepage.asp""     scrolling=""yes"" frameborder=""4"" Title=""Personal Web Site"">"
	Response.Write "</frameset>"
	Response.Write "<noframes>"
	Response.Write "<body>"
	Response.Write "<p>This page uses frames, but your browser doesn't support them.</p>"
	Response.Write "</body>"
	Response.Write "</noframes>"
end if
%>