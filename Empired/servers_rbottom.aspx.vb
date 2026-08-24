Public Class servers_rbottom
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        SwitchPageMode()
    End Sub

    Private Sub SwitchPageMode()
        Dim Mode As String
        Dim Recordset As DataSet, Table As DataTable, Row As DataRow, Column As DataColumn

        Recordset = GetData("sp_Servers")
        DisplayServers(Recordset, 1)
    End Sub

    Private Sub DisplayServers(ByVal Recordset As DataSet, ByVal NoOfTableColumns As Int32)
        Dim Table As DataTable, Row As DataRow
        Dim TableColNumber As Int32
        Dim ServerID As Int32, ServerName As String, ServerIP As String, ServerDescription As String

        ' For each table in the DataSet, output the row values.
        Table = Recordset.Tables(0)
        Response.Write("<table align=""center"" cellspacing=""0"" cellpadding=""0"">" & ControlChars.CrLf)
        Response.Write("    <tr>" & ControlChars.CrLf)
        For Each Row In Table.Rows
            TableColNumber = TableColNumber + 1

            ServerID = Row.Item(0)
            ServerName = Row.Item(1)
            ServerIP = Row.Item(2)
            ServerDescription = Row.Item(3)

            Response.Write("        <td width=""150"" height=""20"" style=""cursor: hand;"">" & ServerName & "</td>" & ControlChars.CrLf)
            Response.Write("        <td width=""100"" height=""20"" style=""cursor: hand;"">" & ServerIP & "</td>" & ControlChars.CrLf)
            Response.Write("        <td width=""200"" height=""20"" style=""cursor: hand;"">" & ServerDescription & "</td>" & ControlChars.CrLf)

            If TableColNumber = NoOfTableColumns Then
                TableColNumber = 0
                Response.Write("    </tr>" & ControlChars.CrLf)
                Response.Write("    <tr>" & ControlChars.CrLf)
            End If
        Next Row
        Response.Write("    </tr>" & ControlChars.CrLf)
        Response.Write("</table>" & ControlChars.CrLf)


    End Sub

End Class
