<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<html>
<%
    Dim i,j
    Dim intTabsWidth
    Dim intTabsCount
    Dim intTabsperRow
    Dim intTabsHeight
    Dim intWidth
    Dim ThisTabWidth
    Dim Divname
    Dim strJavaArray
    Dim intPageHeight
    Dim Tabname
    Dim intTop
    
    intTabsCount = 4
    intTabsperRow = 6
    intTabsWidth = 50
    intTabsHeight = 15
    intPageHeight = 300
    intTop = 0
    Divname ="page"
    intWidth = 50
    Tabname ="tableTab"
    StartX = 90
    StartY = 90
%>
    <head>
        <link rel="stylesheet" type="text/css" href="style.css">
        <title>Tabs</title>
    </head>

<script language='javascript'>
    tabWidth=<%=intTabsWidth%>;
    PageHeight=<%=intPageHeight%>;
    pageWidth=<%=intWidth%>;
<%
    For i = 1 to intTabsCount
        strJavaArray = strJavaArray & chr(34) & Divname & i & chr(34) &  ","
    Next
    strjavaArray = left(strJavaArray,Len(strjavaArray)-1)
    strJavaArray = "var arrPagelayers = new Array(" & strJavaArray & ")"
    response.write "    " & strJavaArray & ";" & vbcrlf
%>
</script>
<script language='javascript'>
    strStored="tableTab1";
    function funTab(layershow) 
    { 
        document.getElementById(window.event.srcElement.id).parentElement.className ="selectedtab";
        if (strStored == "")
        {
    	    strStored = document.getElementById(window.event.srcElement.id).parentElement.id;
    	}
    	if (document.getElementById(window.event.srcElement.id).parentElement.id != strStored) 
        {
    	    document.getElementById(strStored).className = "tabs";
    	    strStored = document.getElementById(window.event.srcElement.id).parentElement.id;
    	}
    	for (i=0; i < arrPagelayers.length; i++) 
        {
    	    document.getElementById(arrPagelayers[i]).style.visibility="hidden";
    	}
        if (document.getElementById(layershow)) 
        {	
            document.getElementById(layershow).style.visibility="visible";
        }
    }

</script>
	<body onload="tableTab1.className='selectedtab';<%=Divname%>1.style.visibility='visible'">
		<div id="tabsmaster" style="position:absolute; left:<%=StartX%>px; top:<%=StartY%>px; width:<%=intTabsWidth%>px; height:<%=intTabsHeight%>px; border-right: thin solid white; visibility: visible;">
<%          For i = 1 to intTabsCount %>
<%              if i = 1 then 
                    ThisTabWidth = 0
                else
                    ThisTabWidth = ThisTabWidth + intWidth
                end if
                j = j + 1
                If j > intTabsperRow Then
                    j = 1  
                    intTop = intTop + intTabsHeight
                    ThisTabWidth = 0
                end if%>
                <div id="tableTab<%=i%>" class="tabs" style="position:absolute; left:<%=ThisTabWidth%>px; top:<%=intTop%>px; width:<%=intWidth%>px; height:<%=intTabsHeight%>px; border-right: thin solid white; visibility: visible;">
                <a href="#" id="pageTab<%=i%>" onclick="funTab('page<%=i%>');">Tab<%=i%></a>
                </div>
<%          Next %>
<%          For i = 1 to intTabsCount%>
                <div id="<%=Divname%><%=i%>" style="position:absolute; left:0px; top:<%=intTop+intTabsHeight%>px; width:300px; height:<%=intPageHeight%>px; border: 2px thin white; visibility: hidden; background-color: steelblue; layer-background-color:steelblue;">
                    page<%=i%>
                </div>
<%          Next %>
		</div>
	</body>
</html>
