<%@ Page Language="vb" AutoEventWireup="false" Codebehind="clients_rbottom.aspx.vb" Inherits="Empired.clients_rbottom"%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>clients_rbottom</title>
		<meta content="Microsoft Visual Studio.NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<link href="Styles.css" type="text/css" rel="stylesheet">
	</HEAD>
	<body ms_positioning="FlowLayout">
		<script language="javascript">
        
        function setheader(ClientID)
        {
        parent.window(1).location = "clients_rtop.aspx?add=false&delete=true";
        parent.window(2).location = "client_details.aspx?clientID=" + ClientID
        }
        
		</script>
		<form id="Form1" method="post" runat="server">
		</form>
	</body>
</HTML>
