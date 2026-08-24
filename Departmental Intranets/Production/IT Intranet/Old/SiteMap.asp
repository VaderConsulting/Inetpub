<!-- #Include file="Includes/Menu.asp" -->

<HTML>
<HEAD>
	<TITLE>IT Intranet: Site Map</TITLE>
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
			<TD width=190><DIV class=HomeLink><A href="index.asp" width=100%>HOME</A></DIV></TD>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Site Map</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<BR>
				<A href="ServiceCentre.asp">Help Desk</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="HelpDeskServices.asp">Help Desk Services</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="UserManuals.asp">User Manuals</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="TipsAndTricks.asp">Tips And Tricks</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="MeetHelpDesk.asp">Meet the Help Desk</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="FAQs.asp">Frequently Asked Questions</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="HotDesk.asp">Hot Desk</A><BR>
				<A href="Procurement.asp">Procurement</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="ProcurementFAQs.asp">Procurement FAQ's</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="ProccurementForm.asp">Procurement Form</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="HardwareList.asp">Hardware List</A><BR>
				<A href="Training.asp">Training</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="TrainManuals.asp">Manuals</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="TrainLiveMeeting.asp">Live Meeting</A><BR>
				<A href="PoliciesAndStandards.asp">Policies and Standards</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="PolicyAcceptableUsage.asp">Acceptable usage policy</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="PolicyRemoteAccess.asp">Remote Access</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="PolicySecurity.asp">Security</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="PolicyProcurement.asp">Procurement</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="PolicyUserAccounts.asp">User Accounts</A><BR>
				<A href="Projects.asp">Projects</A><BR>
				<A href="Forms.asp">Forms</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="ITWorkRequest.asp">IT Work Request Form</A><BR>
				<A href="SelfService.asp">Self Service</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="UpdateDetails.asp">Update My Details</A><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<A href="LogServiceCall.asp">Log a Service Call</A><BR>
				<A href="ContactUs.asp">Contact Us</A><BR>
				<A href="IntranetDisclaimer.asp">Disclaimer</A><BR>
				<A href="IntranetPrivacyStatement.asp">Privacy Statement</A><BR>
				
				<DIV class=Copyright><BR>This page last updated: 2 Aug 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>