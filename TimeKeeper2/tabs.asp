<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">

<html>
<head>
	<link rel="stylesheet" type="text/css" href="style.css">
    <title>Tabs</title>
</head>

<body>
<%
    Dim NoOfTabs, NoOFTabsWide, TabWidth
    Dim i, j
    NoofTabs = 12
    NoOfTabsWide = 8
    TabWidth = 50
%>
<table border='0' cellspacing='0' cellpadding='0' align='left'>
    <tr>
<%
        For i = 1 to NoOfTabs
            j = j + 1
            if j > NoOfTabsWide Then
                j = 0
                Response.Write "</tr>" & vbcrlf
                Response.Write "<tr>" & vbcrlf
            End If
%>
            <td width='<%=TabWidth%>' align='center' nowrap class='tabs' style='border-right: thin solid white;'><a id='tab<%=i%>' name='tab<%=i%>' onclick='TabSelect(<%=i%>);'><span name='span<%=i%>'>Tab <%=i%></span></a></td>
<%
        Next
%>
    </tr>
    <tr>
        <td class='selectedtab' colspan='<%=NoOfTabsWide%>'>&nbsp;</td>
    </tr>
    <tr>
        <td class='selectedtab' colspan='<%=NoOfTabsWide%>'>&nbsp;</td>
    </tr>
    <tr>
        <td class='selectedtab' colspan='<%=NoOfTabsWide%>'>&nbsp;</td>
    </tr>
    <tr>
        <td class='selectedtab' colspan='<%=NoOfTabsWide%>'>&nbsp;</td>
    </tr>
</table>

<script language='JavaScript' defer=true>
    /*
    function TabSelect(TabNo)
    {
        var SelectedTabName;
        var SelectedTab;
        SelectedTabName = "span" + TabNo;
        SelectedTab=document.getElementsByName(SelectedTabName);
        alert (SelectedTab.id);
        //alert (SelectedTab)
    }
    */
</script>

<script language='vbscript'>
    Sub TabSelect(TabNo)
        '' //Dim SelectedTabName
        Dim SelectedTab
        '' //Dim SelectedTabName
        SelectedTabName = "span" & TabNo
        Set SelectedTab=document.getElementsByName(SelectedTabName)
        '' // msgbox SelectedTab
        '' // msgbox TabNo
        SelectedTab.InnerText = "xxx"
    End Sub
</script>
</body>
</html>
