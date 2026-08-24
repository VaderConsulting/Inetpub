<%@ Page Language="vb" AutoEventWireup="false" Codebehind="WebForm1.aspx.vb" Inherits="PWGen.WebForm1"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title></title>
		<meta name="GENERATOR" content="Microsoft Visual Studio.NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK href="Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body MS_POSITIONING="GridLayout">
		<form id="Form1" method="post" runat="server">
			<div align="center">
				<asp:Label id="Label1" runat="server" Width="139px" Height="17px" CssClass="h1">Password Generator</asp:Label>
			</div>
			<div align="right">
				<asp:Label id="Label2" style="Z-INDEX: 101; LEFT: 34px; POSITION: absolute; TOP: 51px" runat="server" Width="232px" Height="15px">Enter the number of passwords required</asp:Label>
			</div>
			<asp:TextBox id="TextBox1" style="Z-INDEX: 102; LEFT: 285px; POSITION: absolute; TOP: 46px" runat="server" Width="70px" Height="21px">1</asp:TextBox>
			<asp:Button id="Button1" style="Z-INDEX: 103; LEFT: 361px; POSITION: absolute; TOP: 42px" runat="server" Width="65px" Height="30px" Text="Generate"></asp:Button>
			<div align="right">
				<asp:Label id="Label3" style="Z-INDEX: 104; LEFT: 90px; POSITION: absolute; TOP: 79px" runat="server" Width="176px" Height="24px">Password Length</asp:Label>
			</div>
			<asp:TextBox id="TextBox2" style="Z-INDEX: 105; LEFT: 285px; POSITION: absolute; TOP: 74px" runat="server" Width="25px" Height="21px">14</asp:TextBox>
			<asp:Button id="Button2" style="Z-INDEX: 106; LEFT: 264px; POSITION: absolute; TOP: 296px" runat="server" Width="57px" Height="21px" Text="Clear"></asp:Button>
			<asp:TextBox id="TextBox3" style="Z-INDEX: 107; LEFT: 89px; POSITION: absolute; TOP: 128px" runat="server" Height="198px" Width="150px" TextMode="MultiLine"></asp:TextBox>
			<div align="right">
			</div>
		</form>
	</body>
</HTML>
