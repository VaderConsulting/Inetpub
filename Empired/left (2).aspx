<%@ Page Language="vb" AutoEventWireup="false" Codebehind="left.aspx.vb" Inherits="Empired.left"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
    <head>
        <title>left</title>
        <meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
        <meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
        <meta name="vs_defaultClientScript" content="JavaScript">
        <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
        <base target="rtop">
        <link href="Styles.css" type="text/css" rel="stylesheet">
    </head>
    <body ms_positioning="FlowLayout">
        <script language="javascript">
            function DoClients()
            {
                parent.window(1).location = "clients_rtop.aspx";
                parent.window(2).location = "clients_rbottom.aspx";
            }
            function DoProjects()
            {
                parent.window(1).location = "projects_rtop.aspx";
                parent.window(2).location = "projects_rbottom.aspx";
            }
            function DoTasks()
            {
                parent.window(1).location = "tasks_rtop.aspx";
                parent.window(2).location = "tasks_rbottom.aspx";
            }
            function DoScheduling()
            {
                parent.window(1).location = "scheduling_rtop.aspx";
                parent.window(2).location = "scheduling_rbottom.aspx";
            }
            function DoServers()
            {
                parent.window(1).location = "servers_rtop.aspx";
                parent.window(2).location = "servers_rbottom.aspx";
            }
            function DoStaff()
            {
                parent.window(1).location = "staff_rtop.aspx";
                parent.window(2).location = "staff_rbottom.aspx";
            }
            function DoHelpdesk()
            {
                parent.window(1).location = "helpdesk_rtop.aspx";
                parent.window(2).location = "helpdesk_rbottom.aspx";
            }
            function DoLocator()
            {
                parent.window(1).location = "locator_rtop.aspx";
                parent.window(2).location = "locator_rbottom.aspx";
            }
            function DoPackaging()
            {
                parent.window(1).location = "packaging_rtop.aspx";
                parent.window(2).location = "packaging_rbottom.aspx";
            }
            function DoPhonebook()
            {
                parent.window(1).location = "phonebook_rtop.aspx";
                parent.window(2).location = "phonebook_rbottom.aspx";
            }
            function DoImages()
            {
                parent.window(1).location = "images_rtop.aspx";
                parent.window(2).location = "images/showimages.asp";
            }
        </script>
        <form id="Form1" method="post" runat="server">
            <table>
                <tr>
                    <td valign="middle"><img height="18" src="images/group.gif" width="18" alt="Clients" border="0" id="IMG1" runat="server" onclick="javascript:DoClients();" style="CURSOR: hand">Clients</A></td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/package.gif" width="18" alt="Projects" border="0" id="IMG2" runat="server" onclick="javascript:DoProjects();" style="CURSOR: hand">Projects</td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/pages.gif" width="18" alt="Tasks" border="0" id="IMG3" runat="server" onclick="javascript:DoTasks();" style="CURSOR: hand">Tasks</td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/schedule.gif" width="18" alt="Scheduling" border="0" id="Img10" runat="server" onclick="javascript:DoScheduling();" style="CURSOR: hand">Scheduling</td>
                </tr>
                <tr>
                    <td valign="middle"><hr></td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/server.gif" width="18" alt="Servers" border="0" id="IMG4" runat="server" onclick="javascript:DoServers();" style="CURSOR: hand">Servers</td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/workstation.gif" width="18" alt="Computers" border="0" id="Img12" runat="server" onclick="javascript:DoComputers();" style="CURSOR: hand">Computers</td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/user2.gif" width="18" alt="Staff" border="0" id="IMG5" runat="server" onclick="javascript:DoStaff();" style="CURSOR: hand">Staff</td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/notes.gif" width="18" alt="Helpdesk" border="0" id="Img11" runat="server" onclick="javascript:DoHelpdesk();" style="CURSOR: hand">Helpdesk</td>
                </tr>
                <tr>
                    <td valign="middle"><hr></td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/earth.gif" width="18" alt="Locator" border="0" id="IMG6" runat="server" onclick="javascript:DoLocator();" style="CURSOR: hand">Locator</td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/addressbook.gif" width="18" alt="Phonebook" border="0" id="IMG7" runat="server" onclick="javascript:DoPhonebook();" style="CURSOR: hand">Phonebook</td>
                </tr>
                <tr>
                    <td valign="middle"><hr></td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/package2.gif" width="18" alt="Packaging" border="0" id="Img9" runat="server" onclick="javascript:DoPackaging();" style="CURSOR: hand">Packaging</td>
                </tr>
                <tr>
                    <td valign="middle"><hr></td>
                </tr>
                <tr>
                    <td valign="middle"><img height="18" src="images/profiles.gif" width="18" alt="Development Images" border="0" id="IMG8" runat="server" onclick="javascript:DoImages();" style="CURSOR: hand">Images</A></td>
                </tr>
                <tr>
                    <td valign="middle"><hr></td>
                </tr>
            </table>
        </form>
    </body>
</html>
