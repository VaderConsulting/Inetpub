<%@ Page Language="vb" AutoEventWireup="false" Codebehind="packaging_rtop.aspx.vb" Inherits="Empired.packaging_rtop"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
    <head>
        <title>packaging_rtop</title>
        <meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
        <meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
        <meta name="vs_defaultClientScript" content="JavaScript">
        <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
        <link href="Styles.css" type="text/css" rel="stylesheet">
    </head>
    <body ms_positioning="GridLayout">
        <script language="javascript">
        
            function doadd()
                {
                parent.window(2).location = "package_details.aspx?mode=new";
                }
            function dodelete()
                {
                parent.window(2).location = "packaging_rbottom.aspx?mode=delete";
                }
            function dostatus()
                {
                parent.window(2).location = "package_status.aspx?mode=new";
                }
            function dotimes()
                {
                parent.window(2).location = "package_timeentry.aspx?mode=new";
                }
            function doall()
                {
                parent.window(1).location = "packaging_rtop.aspx";
                parent.window(2).location = "packaging_rbottom.aspx?mode=all";
                }
            
        </script>
        
        <form id="Form1" method="post" runat="server">
            <h1 align="center">Packaging</h1>
            <table align="center" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="Add" onclick="javascript:doadd();"<%If Delete Then %> disabled<%End If%>>
                    </td>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="Delete" onclick="javascript:dodelete();"<%If Not Delete Then %> disabled<%End If%>>
                    </td>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="Status" onclick="javascript:dostatus();">
                    </td>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="Times" onclick="javascript:dotimes();">
                    </td>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="All" onclick="javascript:doall();">
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
