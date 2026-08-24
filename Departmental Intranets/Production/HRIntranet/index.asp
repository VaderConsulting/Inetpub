<!-- #Include file="Includes/Menu.asp" -->

<html>
<head>
	<title>HR Intranet: Forms</title>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
	<script language=JavaScript>
	<!--
	function HideNonPrintables()
	{
		for (Counter=0;Counter<HidePrint.length;Counter++) HidePrint[Counter].style.display = 'None';
	}
	function ShowNonPrintables()
	{
		for (Counter=0;Counter<HidePrint.length;Counter++) HidePrint[Counter].style.display = 'Inline';
	}
	//-->
	</script>
</head>

<body onbeforeprint="HideNonPrintables()" onafterprint="ShowNonPrintables()">
	<table width=770 cellpadding=0 cellspacing=0 align=center width=770>
		<tr id=HidePrint height=79><td colspan=3 background=Images/Logo.gif align=right class=Copyright><a href="SiteMap.asp">Site Map</a><img align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><a href="ContactUs.asp">Contact Us</a></td></tr>
		<tr id=HidePrint><td colspan=3><img height=11 width=770 src="Images/DottedLine.gif"></td></tr>
		<tr id=HidePrint>
			<td width=190><div class=HomeLink><a onClick="window.close()" width=100%>Back to HR Intranet Home</a></div></td>
			<td width=10>&nbsp;</td>
			<td width=570>&nbsp;</td>
		</tr>
		<tr id=HidePrint><td colspan=3><img height=11 width=770 src="Images/DottedLine.gif"></td></tr>
		<tr>
			<td id=HidePrint width=190 valign=top>
				<table cellpadding=0 cellspacing=0 width=100%>
					<%ShowMenu%>
				</table>
			</td>
			<td width=10></td>
			<td width=570 valign=top>
				<span class=MainHeading style="display:inline;width:460px; height:21px">Forms</span><span class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onClick="window.print()">Print this page</span>
				<p>This section of the intranet will give you access to the Online Human Resources Forms. These forms can be filled out from any Oxiana machine and the request will be emailed directly to the appropriate Department.</p>
				<p class=SubHeading>Please select from the following Forms:</p>
				
				<p><span class=Copyright><a href="<%=NewEmployeeAccessURL%>">New Employee Access</a></span><br>
				Use this form to request network access for a new employee.</p>
				
				<p><span class=Copyright><a href="FormNewNetworkAccess.asp">New Network Access (Existing User)</a></span><br>
				Use this form to request new network access rights for an existing employee.</p>
				
				<p><span class=Copyright><a href="FormExtendAccess.asp">Extend User Access (Existing User)</a></span><br>
				Use this form to extend the duration of network access for an existing Oxiana employee.</p>
				
				<p><span class=Copyright><a href="FormEllipseVaryUser.asp">Ellipse Vary User</a></span><br>
				This form allows you to vary the user permissions within Ellipse for an existing employee.</p>
				
				<p><span class=Copyright><a href="FormNetworkVaryUser.asp">Network Vary User</a></span><br>
				This form allows you to vary the network access for an existing employee.</p>
				
				<p><span class=Copyright><a href="<%=TransferEmployeeURL%>">Transfer Employee Systems Request</a></span><br>
				Use this form when transferring an existing Oxiana employee between different positions.</p>
				
				<p><span class=Copyright><a href="FormTerminatingEmployee.asp">Terminating Employee Access</a></span><br>
				Use this form to disable network access for a terminated Oxiana employee.</p>
				
				<p><span class=Copyright><a href="<%=ChangeOfAddressURL%>">Change of Address</a></span><br>
				Use this form to change your address and contact details.</p>
				
				<div class=Copyright><br>This page last updated: 29 Oct 2007 (Empired)</div>
			</td>
		</tr>
		<tr id=HidePrint><td colspan=3><img height=11 width=770 src="Images/DottedLine.gif"></td></tr>
		<tr id=HidePrint><td colspan=3 class=Copyright><a href="IntranetDisclaimer.asp">Disclaimer</a><img align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><a href="IntranetPrivacyStatement.asp">Privacy</a><img align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</td></tr>
	</TABLE>
</body>
</html>