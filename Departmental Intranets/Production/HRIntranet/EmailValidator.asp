<%@Language = VBScript%>
<%
	response.expires = 0
%>
<html>
<head>
	<TITLE>Email Address Search</TITLE>
	<LINK rel="Stylesheet" type="text/css" href="Styles/Style.css"></LINK>
	<script language="JavaScript">
	function ValidateForm()
	{
		var TheForm = document.forms.frmEmail;
		if (TheForm.FirstName.value.length == 0 && TheForm.LastName.value.length == 0)
		{
			alert('You must specify either a first name or  last name to search for');
			TheForm.FirstName.focus();
			return false;
		}
		window.frames[0].document.body.innerHTML = '<SPAN>Querying the Address book - please be patient...</SPAN>';
	}
	function AddEmail(TheAddress)
	{
		for (var Counter=0;Counter<AllUsers.options.length;Counter++)
		{
			if (AllUsers.options[Counter].value == TheAddress)
			{
				return;
			}
		}
		if (TheAddress.length > 0)
		{
			var NewOption = document.createElement("OPTION");
			NewOption.text = TheAddress;
			NewOption.value = TheAddress;
			AllUsers.options.add(NewOption);
		}
	}
	function SaveAndClose()
	{
		var TempString = '';
		var Counter;
		
		for (Counter=0;Counter<AllUsers.options.length;Counter++)
		{
			if (TempString.length == 0)
			{
				TempString = AllUsers.options[Counter].value;
			}
			else
			{
				TempString += ';' + AllUsers.options[Counter].value;
			}
		}
		
		window.returnValue = TempString;
		window.close();
	}
	function LoadParams()
	{
		var Temp = window.dialogArguments.split(';');
		
		for (var Counter=0;Counter<Temp.length;Counter++)
		{
			AddEmail(Temp[Counter]);
		}
	}
	function RemoveItem()
	{
		if (AllUsers.selectedIndex > -1)
		{
			var Temp = window.confirm('Remove "' + AllUsers.options[AllUsers.selectedIndex].text + '" from your email list?');
			if (Temp)
			{
				AllUsers.options.remove(AllUsers.selectedIndex);
			}
		}
	}
	</script>
</head>

<body onload="LoadParams()">
	<TABLE cellpadding="0" cellspacing="0" width="400">
		<TR height="10" style="FONT-SIZE:1px"><TD>&nbsp;</TD></TR>
		<TR>
			<TD width="10">&nbsp;</TD>
			<TD class="MainHeading">Email Address Search:</td>
			<TD width="10">&nbsp;</TD>
		</tr>
		<tr>
			<TD>&nbsp;</TD>
			<td>
				<form name="frmEmail" action="EmailResults.asp" method="post" target="MatchingUsers" onsubmit="return ValidateForm()">
					<TABLE cellpadding="0" cellspacing="0" width="380">
						<TR>
							<TD>
								<SPAN style="width:90px">First Name:</SPAN><input name="FirstName" style="width:200px" /><br />
								<SPAN style="width:90px">Last Name:</SPAN><input name="LastName" style="width:200px" /><br />
							</TD>
							<TD>
								<input type="submit" value="Search" style="width:75px" />
							</TD>
						</TR>
					</TABLE>
				</form>
			</td>
			<TD>&nbsp;</TD>
		</tr>
		<tr>
			<TD>&nbsp;</TD>
			<td class="SubHeading">
				Search Results:
				<iframe name="MatchingUsers" id="MatchingUsers" style="height:100px;width:380px" src="EmailResults.asp"></iframe>
			</td>
			<TD width="5">&nbsp;</TD>
		</tr>
		<tr><td>&nbsp;</td></tr>
		<tr>
			<TD>&nbsp;</TD>
			<td class="SubHeading">
				Selected Users:
				<select id="AllUsers" size="2" style="height:75px;width:380px" onclick="RemoveItem()">
				</select>
			</td>
			<TD>&nbsp;</TD>
		</tr>
		<TR>
			<TD>&nbsp;</TD>
			<TD align="center"><input type="button" value="OK" onclick="SaveAndClose()" style="Width:75px"> <input type="button" value="Cancel" onclick="window.close()" style="Width:75px"></td>
			<TD>&nbsp;</TD>
		</tr>
	</table>
</body>
</html>