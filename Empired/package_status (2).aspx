<%@ Page Language="vb" AutoEventWireup="false" Codebehind="package_status.aspx.vb" Inherits="Empired.package_status"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>package_details</title>
		<script language="vb" runat="server">

		</script>
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body bgColor="wheat" MS_POSITIONING="GridLayout">
		<script language="vbscript">
            sub SelectPackageName_Click
				PackageStatus.requestType.value = "Get"
				PackageStatus.requestId.value = PackageStatus.SelectPackageName.value
				PackageStatus.submit 
			end sub
		</script>
		<h1>Package Status</h1>
		<form id="PackageStatus" method="post" runat="server">
			<input id="requestType" type="hidden" name="requestType"> <input id="requestId" name="requestId" type="hidden">
			<asp:panel id="panel" Runat="server" Width="509px" Height="89px" BackColor="Wheat">
				<TABLE cellSpacing="0" cellPadding="1" width="80%" bgColor="wheat" border="0">
					<TR>
						<TD>&nbsp;</TD>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD>Package Name: &nbsp;</TD>
						<TD><SELECT id="SelectPackageName" style="WIDTH: 100%" onchange="SelectPackageName_Click" name="SelectPackageName" runat="server"></SELECT></TD>
					</TR>
					<TR>
						<TD>&nbsp;</TD>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD class="SimulateH4" colSpan="2">Status Details:</TD>
					</TR>
					<TR>
						<TD>&nbsp;</TD>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD align="right">Awaiting Software: &nbsp; <INPUT id="CheckAwaitingSoftware" type="checkbox"></TD>
						<TD align="right">Awaiting Hardware: &nbsp; <INPUT id="CheckAwaitingHardware" type="checkbox"></TD>
					</TR>
					<TR>
						<TD align="right">Awaiting Request: &nbsp; <INPUT id="CheckAwaitingRequest" type="checkbox"></TD>
						<TD align="right">Awaiting Personnel: &nbsp; <INPUT id="CheckAwaitingPersonel" type="checkbox"></TD>
					</TR>
					<TR>
						<TD>&nbsp;</TD>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD align="right">Is Building: &nbsp; <INPUT id="CheckIsBuilding" type="checkbox"></TD>
						<TD align="right">Is Testing: &nbsp; <INPUT id="CheckIsTesting" type="checkbox"></TD>
					</TR>
					<TR>
						<TD align="right">Is Validating: &nbsp; <INPUT id="CheckIsValidating" type="checkbox"></TD>
						<TD align="right">Is Complete: &nbsp; <INPUT id="CheckIsComplete" type="checkbox"></TD>
					</TR>
					<TR>
						<TD align="right">Is Published:&nbsp; <INPUT id="CheckIsPublished" type="checkbox"></TD>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD>&nbsp;</TD>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD>&nbsp;</TD>
						<TD>&nbsp;</TD>
					</TR>
					<TR>
						<TD align="middle">
							<INPUT id="ButtonSave" style="WIDTH: 100%" onclick="ButtonSave_Click" type="button" value="Save"></TD>
						<TD align="middle">
							<INPUT id="ButtonRefresh" style="WIDTH: 100%" onclick="ButtonRefresh_Click" type="button" value="Refresh"></TD>
					</TR>
				</TABLE>
			</asp:panel>
			<p>&nbsp;</p>
		</form>
	</body>
</HTML>
