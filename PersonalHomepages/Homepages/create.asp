<%
    ' Create homepage if it doesn't already exist, before redirecting to the (newly created) page.
    ' Always use this page when wishing to view a users homepage.
    'Response.buffer
    Dim ENumber
    ENumber = trim(request.querystring("ENumber"))
    
    If ENumber <> "" Then
        Dim fso, Path, SourcePath, DestPath,NewPath
        Set fso = Server.CreateObject("Scripting.FileSystemObject")

        Path = Request.ServerVariables("PATH_TRANSLATED")
        p = instrrev(Path,"\")
        Path=Left(Path,p)
        SourcePath = replace(Path & "_GTEMPLATE","\","\\")
        DestPath = replace(Path & Enumber,"\","\\")
        
        If fso.FolderExists(DestPath) Then
        Else
        'Copy template
            fso.CopyFolder SourcePath, DestPath, False
        End If
        
        NewPath = Request.ServerVariables("PATH_INFO")
        p = instrrev(NewPath,"/")
        NewPath= Left(NewPath,p) & ENumber & "/default.asp"
        Set fso=Nothing
    Else
        
    End If
%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<head>
	<title>Create Home Page</title>
</head>

<script language="vbscript">
    Sub DoCheck
        <%If ENumber <> "" Then%>
            window.location.href="<%=NewPath%>"
        <%End If%>
    End Sub
</script>
<body onload="DoCheck">
<font size="2" face="Arial, Helvetica, sans-serif" color="#666666">Please wait.  Redirecting to <%=NewPath%>.</font>
</body>
</html>
