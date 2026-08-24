<%@ Page Language="vb" AutoEventWireup="false" Codebehind="package_timeentry.aspx.vb" Inherits="Empired.package_times"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>package_details</title>
		<script language="vb" runat="server">
Private Sub SaveDetails_Click(ByVal sender as System.Object, ByVal e As System.EventArgs)
	If (Not Page.IsValid) then
		return
	end if
	response.write ("Save sucessful")
End Sub

Private sub cancel_Click(ByVal sender as System.Object, ByVal e As System.EventArgs)
	response.write("Cancelled")
End Sub

		</script>
		<meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<h1>Package Time Entries</h1>
		<form id="PackageTimes" method="post" runat="server">
			<asp:panel id="Panel" BackColor="Wheat" Height="89px" Width="709px" Runat="server">
				<P><BR>
					Package Name:
					<asp:DropDownList id="PackageNameDropDown" Height="22px" Width="220px" Runat="server"></asp:DropDownList></P>
				<H2>Packager Details</H2>
				<P>Packager Name:
					<asp:DropDownList id="PackagerNameDropDown" Height="22px" Width="220px" Runat="server"></asp:DropDownList></P>
				<P>Start Time:
					<asp:TextBox id="startTimeTextBox" Height="22px" Width="80px" Runat="server"></asp:TextBox></P>
				<P>&nbsp; End Time:
					<asp:TextBox id="endTimeTextBox" Height="22px" Width="80px" Runat="server"></asp:TextBox></P>
				<asp:RequiredFieldValidator id="validateName" Runat="server" ErrorMessage="Please enter a valid Package Name. " ControlToValidate="PackageNameDropDown"></asp:RequiredFieldValidator>
			</asp:panel>
			<p>
				<asp:Button id="Save" Runat="server" Text="Save Details" OnClick="saveDetails_Click" />
			</p>
			<P>
				<asp:Button id="cancel" Runat="server" Text="Cancel" OnClick="cancel_Click" style="Z-INDEX: 101; LEFT: 156px; POSITION: absolute; TOP: 330px" /></P>
			<p>&nbsp;
			</p>
		</form>
	</body>
</HTML>
