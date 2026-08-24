<%@Language=VBScript%>
<!-- #Include file="Includes/Menu.asp" -->
<!-- #Include file="Includes/FormatDate.asp" -->
<%
	Response.Expires = 0
	
	'check if the user has clicked the submit button
	If Len(Request.Form("SubmitDate")) > 0 Then
		
		'define the local variables
		Dim lvarTemp
		Dim llngTemp
		Dim NumItems
		Dim llngCounter
		Dim lblnSwitch
		lblnSwitch = True
		
		'loop through all of the returned data
		For Each lvarTemp in Request.Form
			
			'check for the HardwareGroup fields
			If Left(lvarTemp, 13) = "HardwareGroup" Then
				
				'find the highest numbered HardwareGroup value - this will indicate how many fields
				'were on the form when submitted
				llngTemp = CLng(Mid(lvarTemp, 14, 15))
				If llngTemp > NumItems Then
					NumItems = llngTemp
				End If
			End If
		Next
		
		
		'Build the HTML body of the email
		Dim lstrFrom
		Dim lstrTo
		Dim lstrCC
		Dim lstrSubject
		Dim lstrHTML
		lstrHTML = ""
		
		'to cc subject
		lstrFrom = Application("EmailProcurementForm")	
		lstrTo = Application("EmailITSupport")	
		lstrSubject = EmailSubjectDev & "Procurement Request Form"
		lstrServer =  Application("EmailSMTPServer")

		lstrHTML = lstrHTML & "<HTML>" & vbcrlf
		lstrHTML = lstrHTML & "<HEAD>" & vbcrlf
		lstrHTML = lstrHTML & "<STYLE>" & vbcrlf
		lstrHTML = lstrHTML & "BODY, TABLE" & vbcrlf
		lstrHTML = lstrHTML & "{" & vbcrlf
		lstrHTML = lstrHTML & "	COLOR:				#555555;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-FAMILY:		Arial, Verdana, Helvetica;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-SIZE:			8pt;" & vbcrlf
		lstrHTML = lstrHTML & "}" & vbcrlf
		lstrHTML = lstrHTML & ".MainHeading" & vbcrlf
		lstrHTML = lstrHTML & "{" & vbcrlf
		lstrHTML = lstrHTML & "	COLOR:				#3C68AE;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-SIZE:			12pt;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-WEIGHT:		Bold;" & vbcrlf
		lstrHTML = lstrHTML & "	HEIGHT:				23px;" & vbcrlf

		lstrHTML = lstrHTML & "}" & vbcrlf
		lstrHTML = lstrHTML & ".SubHeading" & vbcrlf
		lstrHTML = lstrHTML & "{" & vbcrlf
		lstrHTML = lstrHTML & "	COLOR:				#3C68AE;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-SIZE:			11px;" & vbcrlf
		lstrHTML = lstrHTML & "	FONT-WEIGHT:		Bold;" & vbcrlf
		lstrHTML = lstrHTML & "	TEXT-TRANSFORM:		uppercase;" & vbcrlf
		lstrHTML = lstrHTML & "}" & vbcrlf
		lstrHTML = lstrHTML & "</STYLE>" & vbcrlf
		lstrHTML = lstrHTML & "</HEAD>" & vbcrlf
		lstrHTML = lstrHTML & "<BODY>" & vbcrlf
		lstrHTML = lstrHTML & "	<TABLE cellpadding=0 cellspacing=0 width=570>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD class=MainHeading colspan=2><span style=""color:black;"">IT</span>&nbsp;Procurement Request Form</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD width=200>&nbsp;</TD><TD width=370></TD></TR>" & vbcrlf
		
		'now build the HTML to send to the help-desk
		lstrHTML = lstrHTML & "		<TR><TD><B>Submit Date:</B></TD><TD>" & Request.Form("SubmitDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Required By Date:</B></TD><TD>" & Request.Form("RequiredDate") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>User Account:</B></TD><TD>" & Session("SAMUserName") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Name of User:</B></TD><TD>" & Request.Form("User") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Requestor:</B></TD><TD>" & Request.Form("Requestor") & "</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "		<TR><TD><B>Contact Phone:</B></TD><TD>" & Request.Form("ContactPhone") & "</TD></TR>" & vbcrlf
		
		'approval details
		lstrHTML = lstrHTML & "		<TR><TD><B>Site Location:</B></TD><TD>" & Request.Form("Location") & "</TD></TR>" & vbcrlf
		Select Case Request.Form("Location")
			Case ""
				'no fields to display
			Case "Golden Grove"
				'lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
				'lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Approvals</TD></TR>" & vbcrlf
				'lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
				'lstrHTML = lstrHTML & "		<TR><TD><B>Cost Code/Centre:</B></TD><TD>" & Request.Form("CostCode") & "</TD></TR>" & vbcrlf				
				'lstrHTML = lstrHTML & "		<TR><TD><B>Line Manager:</B></TD><TD>" & Request.Form("LineManager") & "</TD></TR>" & vbcrlf
				'lstrHTML = lstrHTML & "		<TR><TD><B>Commercial Manager:</B></TD><TD>" & Request.Form("CommercialManager") & "</TD></TR>" & vbcrlf
				'lstrHTML = lstrHTML & "		<TR><TD><B>General Manager:</B></TD><TD>" & Request.Form("GeneralManager") & "</TD></TR>" & vbcrlf
				
			Case Else
				lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
				lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Approvals</TD></TR>" & vbcrlf
				lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
				lstrHTML = lstrHTML & "		<TR><TD><B>Cost Centre Manager:</B></TD><TD>" & Request.Form("CostCentreManager") & "</TD></TR>" & vbcrlf
				lstrHTML = lstrHTML & "		<TR><TD><B>Cost Code/Centre:</B></TD><TD>" & Request.Form("CostCode") & "</TD></TR>" & vbcrlf
				lstrHTML = lstrHTML & "		<TR><TD><B>Line Manager:</B></TD><TD>" & Request.Form("LineManager") & "</TD></TR>" & vbcrlf
				
		End Select
		lstrHTML = lstrHTML & "		<TR><TD>&nbsp;</TD></TR>" & vbcrlf
		
		'items
		lstrHTML = lstrHTML & "		<TR><TD colspan=2 class=SubHeading>Required Items</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "	</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	<TABLE cellpadding=0 cellspacing=0 width=570>" & vbcrlf
		lstrHTML = lstrHTML & "			<THEAD>" & vbcrlf
		lstrHTML = lstrHTML & "				<TR>" & vbcrlf
		lstrHTML = lstrHTML & "					<TD width=130><B>Item Category</B></TD>" & vbcrlf
		lstrHTML = lstrHTML & "					<TD width=340><B>Item Description</B></TD>" & vbcrlf
		lstrHTML = lstrHTML & "					<TD width=70><B>Price</B></TD>" & vbcrlf
		lstrHTML = lstrHTML & "					<TD width=30><B>Qty</B></TD>" & vbcrlf
		lstrHTML = lstrHTML & "				</TR>" & vbcrlf
		lstrHTML = lstrHTML & "			</THEAD>" & vbcrlf
		lstrHTML = lstrHTML & "			<TBODY id=TableBody>" & vbcrlf
		
		'now add each row (0 to numitems)
		For llngCounter = 0 to NumItems
			lblnSwitch = not lblnSwitch
			if lblnSwitch = true then
				lstrHTML = lstrHTML & "				<TR id=TableRow style=""BACKGROUND-COLOR:#E8E8E8"">" & vbcrlf
			else
				lstrHTML = lstrHTML & "				<TR id=TableRow style=""BACKGROUND-COLOR:#F1F1F1"">" & vbcrlf
			end if
			lstrHTML = lstrHTML & "					<TD>" & Request.form("HardwareGroup" & llngCounter) & "</TD>" & vbcrlf
			If UCase(Request.Form("HardwareGroup" & llngCounter)) = "OTHER..." Then
				lstrHTML = lstrHTML & "					<TD>" & Request.form("txtGroupItem" & llngCounter) & "</TD>" & vbcrlf
			Else
				lstrHTML = lstrHTML & "					<TD>" & Request.form("GroupItem" & llngCounter)& "</TD>" & vbcrlf
			End If
			lstrHTML = lstrHTML & "					<TD>" & Request.form("Price" & llngCounter) & "</TD>" & vbcrlf
			lstrHTML = lstrHTML & "					<TD>" & Request.form("QTY" & llngCounter)& "</TD>" & vbcrlf
			lstrHTML = lstrHTML & "				</TR>" & vbcrlf
		Next
		
		lstrHTML = lstrHTML & "			</TBODY>"
		lstrHTML = lstrHTML & "		</TABLE>"
		lstrHTML = lstrHTML & "		<TABLE cellpadding=0 cellspacing=0 width=570>"
		lstrHTML = lstrHTML & "				<TR><TD width=200>&nbsp;</TD><TD width=370>&nbsp;</TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "				<TR><TD colspan=2><B>Reason For Purchase:</B></TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "				<TR>" & vbcrlf
		lstrHTML = lstrHTML & "					<TD colspan=5>"
		
		'check if a reason was entered
		If Len(Request.Form("Reason")) = 0 Then
			lstrHTML = lstrHTML & "No reason specified"
		Else
			lstrHTML = lstrHTML & Request.Form("Reason")
		End If
		
		lstrHTML = lstrHTML & "					</TD>" & vbcrlf
		lstrHTML = lstrHTML & "				</TR>" & vbcrlf
		lstrHTML = lstrHTML & "				<TR><TD>&nbsp;</TD></TR>"
		lstrHTML = lstrHTML & "				<TR><TD colspan=2><B>Special Instructions to appear on the delivery note:</B></TD></TR>" & vbcrlf
		lstrHTML = lstrHTML & "				<TR>" & vbcrlf
		lstrHTML = lstrHTML & "					<TD colspan=5>"
		
		'check if a reason was entered
		If Len(Request.Form("Special")) = 0 Then
			lstrHTML = lstrHTML & "No special instructions specified"
		Else
			lstrHTML = lstrHTML & Request.Form("Special")
		End If
		
		lstrHTML = lstrHTML & "					</TD>" & vbcrlf
		lstrHTML = lstrHTML & "				</TR>" & vbcrlf
		
		Dim lblnTemp
		on error resume next
		For llngCounter = 0 to NumItems
			lblnTemp = lblnTemp Or (Request.Form("IsMobile" & llngCounter))
		Next
		on error goto 0
		If lblnTemp = true Then
			lstrHTML = lstrHTML & "				<TR><TD>&nbsp;</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "				<TR><TD colspan=2 class=SubHeading>Additional Requirements for Blackberry Requests</TD></TR>" & vbcrlf
			
			if Request.Form("CNumber") = "Existing" Then
				lstrHTML = lstrHTML & "				<TR><TD><B>Mobile Number:</B></TD><TD>Use Existing Number</TD></TR>"
				lstrHTML = lstrHTML & "				<TR><TD>&nbsp;</TD><TD><SPAN style=""width:100px""><B>Mobile Number:</B></SPAN>" & Request.Form("ExistingNumber") & "</TD></TR>" & vbcrlf
				lstrHTML = lstrHTML & "				<TR><TD>&nbsp;</TD><TD><SPAN style=""width:100px""><B>Account Number:</B></SPAN>" & Request.Form("AccountNum") & "</TD></TR>" & vbcrlf
			Else
				lstrHTML = lstrHTML & "				<TR><TD><B>Mobile Number:</B></TD><TD>Assign New Number</TD></TR>" & vbcrlf
			End if
			
			lstrHTML = lstrHTML & "				<TR><TD>&nbsp;</TD></TR>"
			lstrHTML = lstrHTML & "				<TR><TD><B>General Manager:</B></TD><TD>" & Request.Form("GenManager") & "</TD></TR>" & vbcrlf
			lstrHTML = lstrHTML & "				<TR><TD><B>Executive General:</B></TD><TD>" & Request.Form("ExecGeneral") & "</TD></TR>" & vbcrlf
		End If
		lstrHTML = lstrHTML & "			</TBODY>" & vbcrlf
		lstrHTML = lstrHTML & "		</TABLE>" & vbcrlf
		lstrHTML = lstrHTML & "	</BODY>" & vbcrlf
		lstrHTML = lstrHTML & "</HTML>"
		
		'Response.Write lstrHTML
		'Response.End
		
		'Email constants
		Const CdoBodyFormatHTML	= 0
		
		On Error Resume Next
		
		'create the CDO object
		Dim Mail
		Set Mail = Server.CreateObject("CDO.Message")
		
		If Err.number <> 0 Then
			Response.Write "Create CDO.Message Object: " & Err.number & " - " & Err.Description
			Response.End
		End If
		
		'assign the values to the email object
		Mail.To = lstrTo
		Mail.From = lstrFrom
		'Mail.CC = Session("EmailAddr")
		Mail.Subject = lstrSubject
		Mail.HTMLBody = lstrHTML
		
		Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
		Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = lstrServer
		Mail.Configuration.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		Mail.Configuration.Fields.Update()
		
		If Err.number <> 0 Then
			Response.Write "Building the email content: " & Err.number & " - " & Err.Description
			Response.End
		End If
		
		'send the email
		Mail.Send
		
		If Err.number <> 0 Then
			Response.Write "Sending the email: " & Err.number & " - " & Err.Description
			Response.End
		End If
		
		'release the CDO object
		set mail = nothing
		On Error Goto 0
		
		If Err.number <> 0 Then
			Response.Write "Other: " & Err.number & " - " & Err.Description
			Response.End
		End If
%>
<HTML>
<HEAD>
	<TITLE>IT Intranet: Submit Procurement Request</TITLE>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
</HEAD>
<BODY>
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
				<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Submit Procurement Request</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
				<P class=SubHeading>Procurement Request Submitted</P>
				<P>Your request has been emailed to the IT Helpdesk. You will be contacted by one of the Helpdesk staff as soon as possible.</P>
				<DIV class=Copyright><BR>This page last updated: 21 Jul 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>		
<%		
	Else
		
		lstrUserFullName = Session("UsersName")
		lstrEmail = Session("EmailAddr")
%>
<HTML>
<HEAD>
	<TITLE>IT Intranet: Submit Procurement Request</TITLE>
	<link rel="stylesheet" href="Styles/Style.css" type="text/css"></link>
	<SCRIPT language=JavaScript>
	<!--
	var XML;
	var XMLHasLoaded;
	function HideNonPrintables()
	{
		for (Counter=0;Counter<HidePrint.length;Counter++) HidePrint[Counter].style.display = 'None';
	}
	function ShowNonPrintables()
	{
		for (Counter=0;Counter<HidePrint.length;Counter++) HidePrint[Counter].style.display = 'Inline';
	}
	
	function AddRow()
	{
		var Temp = TableBody.innerHTML + '<TR id=TableRow><TD><SELECT name=HardwareGroup' + TableBody.children.length + ' style="WIDTH:130px" onchange="ShowItems(GroupItem' + TableBody.children.length + ', this.options[this.selectedIndex].text)"></SELECT></TD><TD><SELECT name=GroupItem' + TableBody.children.length + ' style="WIDTH:320px" onchange="ShowPrice(Price' + TableBody.children.length + ', IsMobile' + TableBody.children.length + ', HardwareGroup' + TableBody.children.length + '.options[HardwareGroup' + TableBody.children.length + '.selectedIndex].text, this.options[this.selectedIndex].text)"></SELECT><INPUT name=txtGroupItem' + TableBody.children.length + ' style="WIDTH:320px;DISPLAY:None" value="Please enter description here" onfocus="BlankField(this)"></TD><TD><INPUT name=Price' + TableBody.children.length + ' style="WIDTH:70px;BACKGROUND-COLOR:#E8E8E8" readonly><INPUT type=hidden name=IsMobile' + TableBody.children.length + '></TD><TD><INPUT name=QTY' + TableBody.children.length + ' style="WIDTH:50px"></TD></TR>';
		Temp = TableData.innerHTML.substr(0, TableData.innerHTML.indexOf('<TR id=TableRow>')) + Temp + '</TBODY></TABLE>';
		TableData.innerHTML = Temp;
		eval('ShowGroups(Form1.HardwareGroup' + (TableBody.children.length - 1) + ')');
	}
	function LoadXML()
	{
		XML = new ActiveXObject("microsoft.XMLDOM");
		XML.async = false;
		if (XML.load('ITSupportDocuments/XML/Hardware.xml') == true) XMLLoaded();
	}
	function XMLLoaded()
	{
		XMLHasLoaded = true;
		ShowGroups(Form1.HardwareGroup0);
	}
	function ShowGroups(List)
	{
		var Counter;
		if (XMLHasLoaded == true)
		{
			var Groups = XML.getElementsByTagName('GROUP');
			List.add(new Option('-- Select a Group --'));
			for (Counter=0;Counter<Groups.length;Counter++)
			{
				List.add(new Option(Groups[Counter].attributes.getNamedItem('title').nodeValue));
			}
		}
	}
	function ShowItems(List, SelectedGroup)
	{
		var Counter;
		var Counter2;
		if (XMLHasLoaded == true)
		{
			if (SelectedGroup == 'Other...')
			{
				eval('Form1.txt' + List.name + '.style.display = \'inline\'');
				eval('Form1.' + List.name + '.style.display = \'none\'');
			}
			else
			{
				eval('Form1.txt' + List.name + '.style.display = \'none\'');
				eval('Form1.' + List.name + '.style.display = \'inline\'');
				
				var Groups = XML.getElementsByTagName('GROUP');
			
				for (Counter=0;Counter<Groups.length;Counter++)
				{
					if (Groups[Counter].attributes.getNamedItem('title').nodeValue == SelectedGroup)
					{
						var Items = Groups[Counter].getElementsByTagName('ITEM');
						for (Counter2=List.options.length;Counter2>=0;Counter2--) List.options[Counter2] = null;
						
						List.add(new Option('-- Select an Item --'));
						for (Counter2=0;Counter2<Items.length;Counter2++)
						{
							List.add(new Option(Items[Counter2].attributes.getNamedItem('desc').nodeValue));
						}
						return 0;
					}
				}
			}
		}
	}
	function ShowPrice(PriceField, MobileField, SelectedGroup, SelectedItem)
	{
		PriceField.value = '';
		
		var Counter;
		var Counter2;
		var IsMobilePhone = false;
		var MobTemp;
		if (XMLHasLoaded == true)
		{
			var Groups = XML.getElementsByTagName('GROUP');
			
			for (Counter=0;Counter<Groups.length;Counter++)
			{
				if (Groups[Counter].attributes.getNamedItem('title').nodeValue == SelectedGroup)
				{
					var Items = Groups[Counter].getElementsByTagName('ITEM');
					for (Counter2=0;Counter2<Items.length;Counter2++)
					{
						if (Items[Counter2].attributes.getNamedItem('desc').nodeValue == SelectedItem)
						{
							if (Items[Counter2].attributes.getNamedItem('price') != null) PriceField.value = Items[Counter2].attributes.getNamedItem('price').nodeValue;
							if (Items[Counter2].attributes.getNamedItem('mobilephone') != null)
								MobileField.value = Items[Counter2].attributes.getNamedItem('mobilephone').nodeValue;
							else
								MobileField.value = 'false';
						}
					}
				}
			}
		}
		
		for (Counter=0;Counter<(TableBody.children.length);Counter++)
		{
			eval('MobTemp = Form1.IsMobile' + Counter + '.value');
			if (MobTemp == 'True') IsMobilePhone = true;
		}
		if (IsMobilePhone == true)
			MobileOnly.style.display = 'Inline';
		else
			MobileOnly.style.display = 'None';
	}
	function BlankField(TheField)
	{
		if (TheField.value == 'Please enter description here') TheField.value = '';
	}
	function UseExisting(Use)
	{
		if (Use == true)
			TRExisting.style.display = 'inline';
		else
			TRExisting.style.display = 'None'
	}
	function LocationSelect(srcElement)
	{
		switch (srcElement.value)
		{
			case 'Golden Grove':
				ShowHide(GG, true);
				ShowHide(NONGG, false);
				ShowHide(NONE, true);
				break;
			case  '':
				ShowHide(GG, false);
				ShowHide(NONGG, false);
				ShowHide(NONE, false);
				break;
			default:
				ShowHide(GG, false);
				ShowHide(NONGG, true);
				ShowHide(NONE, true);
				break;
		}
	}
	function ShowHide(Elements, Show)
	{
		var Counter;
		for (Counter=0;Counter<Elements.length;Counter++)
		{
			if (Show == false)
			{
				Elements[Counter].style.display = 'None';
			}
			else
			{
				Elements[Counter].style.display = 'Inline';
			}
		}
	}
	function ValidateAndSubmit()
	{
		var TheForm = document.forms.Form1;
		var Counter;
		var Temp;
		var IsMobile = false;
		
		//validate the fields
		if (CheckField(TheForm.RequiredDate, '', 'The Required By date must be entered.') == false) return false;
		if (CheckField(TheForm.User, '', 'Please specify the user who the items are for..') == false) return false;
		if (CheckField(TheForm.ContactPhone, '', 'Please specify a contact phone number') == false) return false;
		if (CheckField(TheForm.Location, '', 'Please specify your location.') == false) return false;
		
		//check if the site is not golden grove and validate the cost centre fields
		if (TheForm.Location.value != 'Golden Grove')
		{
			if (CheckField(TheForm.CostCentreManager, '', 'You must specify a cost centre manager for this request.') == false) return false;
			if (CheckField(TheForm.CostCode, '', 'You must specify a cost code for this request.') == false) return false;
			if (CheckField(TheForm.LineManager, '', 'You must specify your Line Manager for this request.') == false) return false;
		}
		
		//only one row of items has been added, so validate it
		if (TableRow.length == null)
		{
			
			//check if the user has selected a group
			if (TheForm.HardwareGroup0.item(TheForm.HardwareGroup0.selectedIndex).text != '-- Select a Group --')
			{
				//check that the user has specified an item
				if (TheForm.HardwareGroup0.item(TheForm.HardwareGroup0.selectedIndex).text == 'Other...')
				{
					if (CheckField(TheForm.txtGroupItem0, '', 'Please specify the Item you require.') == false) return false;
					if (CheckField(TheForm.txtGroupItem0, 'Please enter description here', 'Please specify the Item you require.') == false) return false;
				}
				else
				{
					if (TheForm.GroupItem0.item(TheForm.GroupItem0.selectedIndex).text == '-- Select an Item --')
					{
						alert('Please select the Item you wish to request.');
						TheForm.GroupItem0.focus();
						return false;
					}
				}
				
				//check the qty field and determine if the item is a mobile
				if (CheckField(TheForm.QTY0, '', 'Please specify the number of Items required') == false) return false;
				if (TheForm.IsMobile0.value == 'True') IsMobile = true;
			}
		}
		
		//the user has added multiple Items to the form, so validate them all
		else
		{
			
			//loop through all of the items
			for (Counter=0;Counter<TableRow.length;Counter++)
			{
				
				//check that the user has specified a group for this item
				eval('Temp = document.forms.Form1.HardwareGroup' + Counter);
				if (Temp.item(Temp.selectedIndex).text != '-- Select a Group --')
				{
					//check that the user has specified an Item
					if (Temp.item(Temp.selectedIndex).text == 'Other...')
					{
						eval('Temp = TheForm.txtGroupItem' + Counter);
						if (CheckField(Temp, '', 'Please specify the Item you require.') == false) return false;
						if (CheckField(Temp, 'Please enter description here', 'Please specify the Item you require.') == false) return false;
					}
					else
					{
						eval('Temp = TheForm.GroupItem' + Counter);
						if (Temp.item(Temp.selectedIndex).text == '-- Select an Item --')
						{
							alert('Please select the Item you wish to request.');
							Temp.focus();
							return false;
						}
					}
					//check the qty field and determine if the item uis a mobile
					eval('Temp = TheForm.QTY' + Counter);
					if (CheckField(Temp, '', 'Please specify the number of Items required') == false) return false;
					eval('Temp = TheForm.IsMobile' + Counter);
					if (Temp != null) if (Temp.value == 'True') IsMobile = true;
				}
			}
		}
		
		//check that the user has specified a reason
		if (CheckField(TheForm.Reason, '', 'Please specify your reason for this procurement request.') == false) return false;
		
		//if one or more of the Items was a mobile device, validate the extra fields
		if (IsMobile == true)
		{
			//validate the extra fields
			if ((TheForm.CNumber[0].checked == false) && (TheForm.CNumber[1].checked == false))
			{
				alert('Plase specify if you would like to use a new or existing number for the mobile device.');
				TheForm.CNumber[0].focus();
				return false;
			}
			if (TheForm.CNumber[1].checked) if (CheckField(TheForm.ExistingNumber, '', 'You must specify your existing number.') == false) return false;
			if (CheckField(TheForm.GenManager, '', 'You must specify your General Manager when requesting a mobile device.') == false) return false;
			if (CheckField(TheForm.ExecGeneral, '', 'You must specify your Executive General Manager when requesting a mobile device.') == false) return false;
		}
		TheForm.submit();
	}
	function CheckField(TheField, BadValue, Message)
	{
		if (TheField.value == BadValue)
		{
			alert(Message);
			TheField.focus();
			return false;
		}
	}
	//-->
	</SCRIPT>
	<SCRIPT language=javascript src="Scripts/HTMLDateFormat.js"></SCRIPT>
</HEAD>

<BODY onload="LoadXML()" onbeforeprint="HideNonPrintables()" onafterprint="ShowNonPrintables()">
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
				<FORM id=Form1 method=post action="ProcurementForm.asp">
					<TABLE cellpadding=0 cellspacing=0 width=570 id=form>
						<TR>
							<TD colspan=2>
								<SPAN class=MainHeading style="display:inline;width:460px; height:21px">Submit Procurement Request</SPAN><SPAN class=Copyright style="padding-top:6px;padding-right:25px;display:inline;width:110px;HEIGHT:21px;cursor:hand;BACKGROUND-IMAGE:url('Images/PrintArrow.gif')" onclick="window.print()">Print this page</SPAN>
							</TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD width=200><B>Submit Date:</B></TD>
							<TD width=370><INPUT name=SubmitDate value="<%= FormatDate(Now, C_FormatDate_ShortDate)%>" onblur="CheckDate(this, true)" style="WIDTH:75px;BACKGROUND-COLOR:#E8E8E8" readonly></TD>
						</TR>
						<TR>
							<TD><B>Required By Date: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=RequiredDate onblur="CheckDate(this, true)" style="WIDTH:75px"></TD>
						</TR>
						<TR>
							<TD colspan=2>(Depending on stock availability, please allow at least 7 days for delivery from Management approval)</TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>Name of User: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=User></TD>
						</TR>
						<TR>
							<TD><B>Requestor:</B></TD>
							<TD><INPUT name=Requestor value="<%= lstrUserFullName%>" readonly style="BACKGROUND-COLOR:#E8E8E8"></TD>
						</TR>
						<TR>
							<TD><B>Contact Phone: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=ContactPhone></TD>
						</TR>
						<TR>
							<TD><B>Site Location: <SPAN class=Warning>*</SPAN></B><BR>(where item will be located)</TD>
							<TD>
								<SELECT name=Location onchange="LocationSelect(this)">
									<OPTION value="">-- Select a Location --</OPTION>
									<OPTION value="Golden Grove">Golden Grove</OPTION>
									<OPTION value="Melbourne">Melbourne</OPTION>
									<OPTION value="Adelaide">Adelaide</OPTION>
									<OPTION value="SPDG">SPDG</OPTION>
									<OPTION value="Prominent Hill">Prominent Hill</OPTION>
									<OPTION value="Other">Other...</OPTION>
								</SELECT>
							</TD>
						</TR>
						
						<TR id=NONE style="DISPLAY:None"><TD>&nbsp;</TD></TR>
						<TR id=NONE style="DISPLAY:None"><TD colspan=2 class=SubHeading>Approvals</TD></TR>
						<TR id=NONE style="DISPLAY:None"><TD>&nbsp;</TD></TR>
						<TR id=NONGG style="DISPLAY:None">
							<TD><B>Cost Centre Manager: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=CostCentreManager></TD>
						</TR>
						<TR id=NONGG style="DISPLAY:None">
							<TD><B>Cost Code/Centre: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=CostCode title="Account for purchase to be charged to"></TD>
						</TR>
						<TR id=NONGG style="DISPLAY:None">
							<TD><B>Line Manager: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=LineManager></TD>
						</TR>
						<TR id=GG style="DISPLAY:None">
							<TD id=GG colspan=2>Prior to the purchase of any IT Equipment (except Blackberries), the current approval process, including submission of the minor capital AFE will need to be followed</TD>
						</TR>
						<!--<TR id=GG style="DISPLAY:None">
							<TD><B>General Manager: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=GeneralManager></TD>
						</TR>-->
						
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Required Items</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD colspan=2 id=TableData>
								<TABLE cellpadding=0 cellspacing=0 width=570>
									<THEAD>
										<TR>
											<TD width=130><B>Item Category <SPAN class=Warning>*</SPAN></B></TD>
											<TD width=320><B>Item Description <SPAN class=Warning>*</SPAN></B></TD>
											<TD width=70><B>Price</B></TD>
											<TD width=50><B>Qty <SPAN class=Warning>*</SPAN></B></TD>
										</TR>
									</THEAD>
									<TBODY id=TableBody>
										<TR id=TableRow>
											<TD><SELECT name=HardwareGroup0 style="WIDTH:130px" onchange="ShowItems(GroupItem0, this.options[this.selectedIndex].text)"></SELECT></TD>
											<TD><SELECT name=GroupItem0 style="WIDTH:320px" onchange="ShowPrice(Price0, IsMobile0, HardwareGroup0.options[HardwareGroup0.selectedIndex].text, this.options[this.selectedIndex].text)"></SELECT><INPUT name=txtGroupItem0 style="WIDTH:320px;DISPLAY:None" value="Please enter description here" onfocus="BlankField(this)"></TD>
											<TD><INPUT name=Price0 style="WIDTH:70px;BACKGROUND-COLOR:#E8E8E8" readonly><INPUT type=hidden name=IsMobile0></TD>
											<TD><INPUT name=QTY0 style="WIDTH:50px"></TD>
										</TR>
									</TBODY>
								</TABLE>
							</TD>
						</TR>
						<TR><TD><SPAN class=Copyright><A href="#" onclick="AddRow();return false;">Add Another</A></SPAN></TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Reason for purchase: <SPAN class=Warning>*</SPAN><B></TD></TR>
						<TR>
							<TD colspan=2><TEXTAREA name=Reason style="HEIGHT:50px;WIDTH:570px"></TEXTAREA></TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Optional</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2><B>Special Instructions to appear on the delivery note:<B></TD></TR>
						<TR>
							<TD colspan=2><TEXTAREA name=Special style="HEIGHT:50px;WIDTH:570px"></TEXTAREA></TD>
						</TR>
					</TABLE>
					<TABLE cellpadding=0 cellspacing=0 width=570 id=MobileOnly style="DISPLAY:None">
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD colspan=2 class=SubHeading>Additional Requirements for Blackberry Requests</TD></TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD width=200><B>Mobile Number: <SPAN class=Warning>*</SPAN></B></TD>
							<TD width=370><INPUT id=NewNum type=radio name=CNumber value=New onclick="UseExisting(false)">Assign New Number&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT type=radio name=CNumber value=Existing onclick="UseExisting(true)">Use Existing Number</TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR id=TRExisting style="display:none">
							<TD>&nbsp;</TD>
							<TD>When requesting the transfer of an existing mobile number, you will need to:<BR>
								<LI>Print and fill out a <SPAN class=Copyright><A href="<%=Application("Documents")%>Forms/Procurement/ConsumerTransferofService.pdf" target="Transfer">Transfer of Ownership</A></SPAN> form and fax it to Melbourne IT on +61 3 8623 2233)
								<LI>Provide your existing mobile account number (located on the front of your mobile phone account).<BR>
								<SPAN style="WIDTH:150px"><B>Mobile Number: <SPAN class=Warning>*</SPAN></B></SPAN><INPUT name=ExistingNumber style="WIDTH:100px"><BR>
								<SPAN style="WIDTH:150px"><B>Account Number:</B></SPAN><INPUT name=AccountNum style="WIDTH:100px">
							</TD>
						</TR>
						<TR><TD>&nbsp;</TD></TR>
						<TR>
							<TD><B>General Manager: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=GenManager></TD>
						</TR>
						<TR>
							<TD><B>Executive General Manager: <SPAN class=Warning>*</SPAN></B></TD>
							<TD><INPUT name=ExecGeneral></TD>
						</TR>
					</TABLE>
					<TABLE cellpadding=0 cellspacing=0 width=570>
						<TR><TD>&nbsp;</TD></TR>
						<TR><TD align=center><INPUT type=button value=Submit onclick="ValidateAndSubmit()"></TD></TR>
					</TABLE>
				</FORM>
				<DIV class=Copyright><BR>This page last updated: 10 Sep 2006 (Empired)</DIV>
			</TD>
		</TR>
		<TR id=HidePrint><TD colspan=3><IMG height=11 width=770 src="Images/DottedLine.gif"></TD></TR>
		<TR id=HidePrint><TD colspan=3 class=Copyright><A href="IntranetDisclaimer.asp">Disclaimer</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif"><A href="IntranetPrivacyStatement.asp">Privacy</A><IMG align="absmiddle" width=19 height=12 src="Images/PrivacyDivider.gif">Copyright Oxiana Limited 2006</TD></TR>
	</TABLE>
</BODY>
</HTML>
<%
	End If
%>