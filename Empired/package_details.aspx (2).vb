Public Class package_details
    Inherits System.Web.UI.Page
    Protected WithEvents PackageNameTextBox As System.Web.UI.WebControls.TextBox
    Protected WithEvents PackageStatusTextBox As System.Web.UI.WebControls.TextBox
    Protected WithEvents validateName As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents validateStatus As System.Web.UI.WebControls.RequiredFieldValidator
    Protected WithEvents validatePackage As System.Web.UI.WebControls.CustomValidator
    Protected WithEvents panel As System.Web.UI.WebControls.Panel
    Protected WithEvents Save As System.Web.UI.WebControls.Button
    Protected WithEvents cancel As System.Web.UI.WebControls.Button

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
        If Not Page.IsPostBack Then
            Dim Recordset As DataSet
            Recordset = GetData("sp_AllPackageInfo")
            '            PackageNameTextBox.datasource = Recordset
            '           PackageNameTextBox.DataBind()
        End If
    End Sub

    Private Sub SwitchPageMode()
        Dim Mode As String
        Dim Table As DataTable, Row As DataRow, Column As DataColumn

        Mode = Page.Request.QueryString("Mode")

        Select Case Mode
            Case "new"
                defaultDetails()
            Case "edit"
                plantDetails()
        End Select

    End Sub

    Private Sub defaultDetails()

    End Sub

    Private Sub plantDetails()

    End Sub

End Class
