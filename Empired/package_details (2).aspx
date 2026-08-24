<%@ Page Language="vb" AutoEventWireup="false" Codebehind="package_details.aspx.vb" Inherits="Empired.package_details"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>package_details</title>
		<script language="vb" runat="server">
Protected Sub validatePackageDetails (source as Object, args As System.Web.UI.WebControls.ServerValidateEventArgs)
	response.write("validated")
	args.IsValid = true
End Sub

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
		<h1>Package Details</h1>
		<form id="PackageDetails" method="post" runat="server">
			<asp:Panel id="panel" Runat="server" Width="504px" Height="89px" BackColor="Wheat"><BR>Package Name: 
<asp:TextBox id="PackageNameTextBox" Height="22px" Width="80px" Runat="server"></asp:TextBox>Package 
Status: 
<asp:TextBox id="PackageStatusTextBox" Height="22px" Width="80px" Runat="server"></asp:TextBox><BR>
<asp:RequiredFieldValidator id="validateName" Runat="server" ControlToValidate="PackageNameTextBox" ErrorMessage="Please enter a valid Package Name. "></asp:RequiredFieldValidator>
<asp:RequiredFieldValidator id="validateStatus" Runat="server" ControlToValidate="PackageStatusTextBox" ErrorMessage="Please enter A valid Status. "></asp:RequiredFieldValidator>
<asp:CustomValidator id="validatePackage" Runat="server" ControlToValidate="PackageStatusTextBox" OnServerValidate="validatePackageDetails"></asp:CustomValidator>
		</asp:Panel>
			<p>
				<asp:Button id="Save" Runat="server" Text="Save Details" OnClick="saveDetails_Click" />
			</p>
			<p>
				<asp:Button id="cancel" Runat="server" Text="Cancel" OnClick="cancel_Click" />
			</p>
		</form>
	</body>
</HTML>
