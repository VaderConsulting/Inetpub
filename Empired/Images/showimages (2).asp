<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
    <head>
        <title>Image List</title><link href="../Styles.css" type="text/css" rel="stylesheet">
    </head>
    <body>
        <table border="1" cellspacing="0" cellpadding="0">
            <%
            Set FileSystem = CreateObject("Scripting.FileSystemObject")
            Path = Request.ServerVariables ("PATH_TRANSLATED")
            j = InStrRev(Path,"\")
            Path = Left(Path,j-1)

            Set oFolder = FileSystem.GetFolder(Path)
            Set oFiles = oFolder.Files
            For Each oF in oFiles 
                If right(oF.Name,3) <> "asp" Then%>
                    <tr>
                        <td><%= oF.Name %></td>
                        <td><IMG SRC="<%= oF.Name %>"></td>
                    </tr>
                <%End If
             Next %>
        </table>
    </body>
</html>
