<!-- #Include file="Includes/Menu.asp" -->

<HTML>
<HEAD>
	<TITLE>IT Intranet: Hot Desk</TITLE>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Booking an Oxiana Perth Hot Desk</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P>Access to the Oxiana computer network is provided in Perth, Western Australia by Empired Ltd located at 469 Murray St. Perth WA.  Facilities are available 6am to 6pm WST</P>
				<DIV class=SubHeading>Facilities available</DIV>
				<UL>
					<LI>2 Hot Desks with Oxiana desktop computers (you can also connect your laptop to Oxiana at these desks)</LI>
					<LI>Tea and Coffee making facilities</LI>
					<LI>Meeting Room (8 seats)</LI>
					<LI>Training Room with Projector (available on request)</LI>
					<LI>Kitchen, Toilets and Shower Facilities</LI>
				</UL>
				<DIV class=SubHeading>Bookings</DIV>
				<OL>
					<LI>Book the Hot Desk or Meeting room using an Outlook Meeting Request (They are named “Resource - Perth xxxxx” eg. HD1, HD2, Meeting Room). The step by step guide to booking resources in Outlook can be found on page 2 and 3 of this document.</LI>
					<LI>Within 24 hours, Monday to Friday, Empired Reception will accept your meeting request electronically.</LI>
					<LI>Optional: You may wish to verbally book a room by calling Empired on 08 9321 9401 should you not have access to the Outlook booking method.</LI>
					<LI>When arriving at Empired please report to reception. There will be a register available that you will be required to sign both in and out of.  An Oxiana identity badge will also be provided.</LI>
					<LI>There is an information sheet located near the hot desks detailing Emergency Exits, Fire Extinguishers, Toilets, and Kitchen etc</LI>
					<LI>If you require anything else while at Empired please ask reception.</LI>
				</OL>
				<DIV class=SubHeading>Finding Empired Ltd (469 Murray St)</DIV><BR>
				&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<IMG height=306 width=518 src="Images/HotDesk/Location.gif">
				<P>If you have any further queries or wish to request additional facilities please contact Empired on 08 9321 9401</P>
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Booking Resources using Outlook</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P>Open Outlook click the Calendar button on the lower left of the screen.</P>
				<P>Next click the <IMG align="absmiddle" height=22 width=63 src="Images/HotDesk/New.gif"> button from the toolbar.</P>
				<IMG height=166 width=570 src="Images/HotDesk/Appointment.gif">
				<P>Fill in the Subject, Time and Location.</P>
				<P>Next click the <B>Scheduling</B> tab.</P>
				<IMG height=370 width=570 src="Images/HotDesk/Scheduling.gif">
				<P>At the bottom of the Scheduling tab click the <IMG align="absmiddle" height=23 width=99 src="Images/HotDesk/AddOthers.gif"> Button. (Highlighted in Green)</P>
				<IMG height=418 width=518 src="Images/HotDesk/Attendees.gif">
				<P>You can select people from the Global Address List required at the meeting and add them by Clicking the <IMG align="absmiddle" height=23 width=86 src="Images/HotDesk/Required.gif"> Button.</P>
				<P>Find the name of the resource from the global address list and add it <IMG align="absmiddle" height=25 width=82 src="Images/HotDesk/Resources.gif"> Click <B>OK</B>.</P>
				<P>To create the booking / appointment click <B>send</B> from the toolbar.</P>
				<P>The booking will be automatically forwarded to Empired Ltd reception who will accept or decline depending on resource availability.</P>
				<IMG height=183 width=570 src="Images/HotDesk/Scheduling2.gif">
				<DIV class=Copyright><BR>This page last updated: 07 Jul 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>