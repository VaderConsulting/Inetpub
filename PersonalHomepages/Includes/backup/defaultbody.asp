<%@ Language=VBScript %>
<%
option explicit
Response.Buffer = True
'On Error Resume Next
%>
<!--#include file="constants.asp"-->
<!--#include file="whologgedin.asp"-->
<%
if lcase(UserNameReturn) = lcase(PathNameReturn) then
	Response.Write "<frameset framespacing=""1"" border=""0"" frameborder=""NO"" rows=""20%,80%"" bordercolor=""#FFFFFF"">"
	Response.Write "	<frame name=""frmMain"" src=""homepage.htm""     scrolling=""yes"" noresize frameborder=""NO"" Title=""Personal Web Site"">"
	Response.Write "	<frame name=""frmEdit"" src=""../../includes/edithomepage.asp"" scrolling=""yes"" noresize frameborder=""NO"" Title=""Edit Personal Web Site"">"
	Response.Write "</frameset>"
	Response.Write "<noframes>"
	Response.Write "<body>"
	Response.Write "<p>This page uses frames, but your browser doesn't support them.</p>"
	Response.Write "</body>"
	Response.Write "</noframes>"
else
	Response.Write "<frameset framespacing=""1"" border=""0"" frameborder=""NO"" rows=""100%"" bordercolor=""#FFFFFF"">"
	Response.Write "	<frame name=""frmMain"" src=""homepage.htm""     scrolling=""yes"" noresize frameborder=""NO"" Title=""Personal Web Site"">"
	Response.Write "</frameset>"
	Response.Write "<noframes>"
	Response.Write "<body>"
	Response.Write "<p>This page uses frames, but your browser doesn't support them.</p>"
	Response.Write "</body>"
	Response.Write "</noframes>"
end if
%>