<%@ Page Language="vb" AutoEventWireup="false" Codebehind="clients_rtop.aspx.vb" Inherits="Empired.clients"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
    <head>
        <title>clients</title>
        <meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
        <meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
        <meta name="vs_defaultClientScript" content="JavaScript">
        <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
        <link href="Styles.css" type="text/css" rel="stylesheet">
    </head>
    <body ms_positioning="FlowLayout">
        <script language="javascript">
        
            function doadd()
                {
                parent.window(2).location = "clients_rbottom.aspx?mode=add";
                }
            function dodelete()
                {
                parent.window(2).location = "clients_rbottom.aspx?mode=delete";
                }
            function doall()
                {
                parent.window(1).location = "clients_rtop.aspx";
                parent.window(2).location = "clients_rbottom.aspx?mode=all";
                }
            function docurrent()
                {
                parent.window(1).location = "clients_rtop.aspx";
                parent.window(2).location = "clients_rbottom.aspx?mode=current";
                }
            
        </script>
        <form id="Form1" method="post" runat="server">
            <h1 align="center">Clients</h1>
            <table align="center" border="0" cellpadding="0" cellspacing="0">
                <tr>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="Add" id="btn_Add" onclick="javascript:doadd();"<%If Not Add Then %> disabled<%End If%>>
                    </td>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="Delete" onclick="javascript:dodelete();"<%If Not Delete Then %> disabled<%End If%>>
                    </td>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="Show All" onclick="javascript:doall();">
                    </td>
                    <td width="120" align="middle">
                        <input style="width: 100px" type="button" value="Show Current" onclick="javascript:docurrent();">
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
