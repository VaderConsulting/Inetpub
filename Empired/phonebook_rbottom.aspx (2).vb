Public Class phonebook_rbottom
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
        Dim ErrorMsg As String

        Recordset = GetData("sp_Users", ErrorMsg)
        If ErrorMsg = "" Then
            DisplayData(Recordset, 1)
        End If
    End Sub

    Private Sub DisplayData(ByVal Recordset As DataSet, ByVal NoOfTableColumns As Int32)
        Dim Table As DataTable, Row As DataRow
        Dim TableColNumber As Int32
        Dim ID As Int32, CName As String, SName As String, Mobile As String, Extension As String

        ' For each table in the DataSet, output the row values.
        Table = Recordset.Tables(0)
        Response.Write("<table align=""center"" cellspacing=""0"" cellpadding=""0"">" & ControlChars.CrLf)
        Response.Write("    <tr>" & ControlChars.CrLf)
        For Each Row In Table.Rows
            TableColNumber = TableColNumber + 1

            ID = Row.Item(0)
            CName = Row.Item(1) & ""
            SName = Row.Item(2) & ""
            Mobile = Row.Item(4) & ""
            Extension = Row.Item(5) & ""

            ' Turn number into form of xxxx xxx xxx
            If Mobile <> "" Then
                Mobile = Mid(Mobile, 1, 4) & " " & Mid(Mobile, 5, 3) & " " & Mid(Mobile, 8, 3)
            End If

            Response.Write("        <td width=""80"" height=""20"" style=""cursor: hand;"">" & CName & "</td>" & ControlChars.CrLf)
            Response.Write("        <td width=""100"" height=""20"" style=""cursor: hand;"">" & SName & "</td>" & ControlChars.CrLf)
            Response.Write("        <td width=""100"" height=""20"" style=""cursor: hand;"">" & Mobile & "</td>" & ControlChars.CrLf)
            Response.Write("        <td width=""100"" height=""20"" style=""cursor: hand;"">" & Extension & "</td>" & ControlChars.CrLf)

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
