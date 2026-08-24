<!-- #Include file="Includes/Menu.asp" -->

<HTML>
<HEAD>
	<TITLE>HR Intranet: Forms</TITLE>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
	<SCRIPT language=JavaScript>
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
	</SCRIPT>
</HEAD>

<BODY onbeforeprint="HideNonPrintables()" onafterprint="ShowNonPrintables()">
	<TABLE width=770 cellpadding=0 cellspacing=0 align=center width=770>
		<TR id=HidePrint height=79><TD colspan=3 background=Images/Logo.gif align=right class=Copyright><A href="SiteMap.asp">Site Map</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="ContactUs.asp">Contact Us</A></TD></TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint>
			<TD width=190><DIV class=HomeLink><A onclick="window.close()" width=100%>Back to HR Intranet Home</A></DIV></TD>
			<TD width=10>&nbsp;</TD>
			<TD width=570>&nbsp;</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR>
			<TD id=HidePrint width=190 valign=top>
				<TABLE cellpadding=0 cellspacing=0 width=100%>
					<%ShowMenu%>
				</TABLE>
			</TD>
			<TD width=10></TD>
			<TD width=570 valign=top>
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Forms</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P>This section of the intranet will give you access to the Online Human Resources Forms. These forms can be filled out from any Oxiana machine and the request will be emailed directly to the appropriate Department.</P>
				<P class=SubHeading>Please select from the following Forms:</P>
				
				<P><SPAN class=Copyright><A href="FormNewEmployee.asp">New Employee Access</A></SPAN><BR>
				Use this form to request network access for a new employee.</P>
				
				<P><SPAN class=Copyright><A href="FormNewNetworkAccess.asp">New Network Access (Existing User)</A></SPAN><BR>
				Use this form to request new network access rights for an existing employee.</P>
				
				<P><SPAN class=Copyright><A href="FormExtendAccess.asp">Extend User Access (Existing User)</A></SPAN><BR>
				Use this form to extend the duration of network access for an existing Oxiana employee.</P>
				
				<P><SPAN class=Copyright><A href="FormEllipseVaryUser.asp">Ellipse Vary User</A></SPAN><BR>
				This form allows you to vary the user permissions within Ellipse for an existing employee.</P>
				
				<P><SPAN class=Copyright><A href="FormNewtworkVaryUser.asp">Network Vary User</A></SPAN><BR>
				This form allows you to vary the network access for an existing employee.</P>
				
				<P><SPAN class=Copyright><A href="FormTransferEmployee.asp">Transfer Employee Systems Request</A></SPAN><BR>
				Use this form when transferring an existing Oxiana employee between different positions.</P>
				
				<P><SPAN class=Copyright><A href="FormTerminateEmployee.asp">Terminating Employee Access</A></SPAN><BR>
				Use this form to disable network access for a terminated Oxiana employee.</P>
				
				<P><SPAN class=Copyright><A href="FormChangeOfAddress.asp">Change of Address</A></SPAN><BR>
				Use this form to change your address and contact details.</P>
				
				<DIV class=Copyright><BR>This page last updated: 18 Dec 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>