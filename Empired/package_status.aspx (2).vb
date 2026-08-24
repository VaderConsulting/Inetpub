Public Class package_status
    Inherits System.Web.UI.Page
    Protected WithEvents panel As System.Web.UI.WebControls.Panel
    Protected WithEvents Save As System.Web.UI.WebControls.Button
    Protected WithEvents CheckAwaitingSoftware As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents CheckAwaitingHardware As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents CheckAwaitingRequest As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents CheckAwaitingPersonel As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents CheckIsBuilding As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents CheckIsTesting As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents CheckIsValidating As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents CheckIsComplete As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents CheckIsPublished As System.Web.UI.HtmlControls.HtmlInputCheckBox
    Protected WithEvents ButtonSave As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents ButtonCancel As System.Web.UI.HtmlControls.HtmlInputButton
    Protected WithEvents SelectPackageName As System.Web.UI.HtmlControls.HtmlSelect
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
        Dim requestType As String
        Dim requestId As String

        requestType = Page.Request.Form("RequestType")
        requestId = Page.Request.Form("RequestId")

        If requestType <> "" Then
            Response.Write(requestType)
            Response.Write(requestId)
            Response.Flush()
            Select Case requestType
                Case "Get"
                    getInfoFor(requestId)
            End Select
        Else
            LoadDropDown()
        End If
    End Sub

    Private Sub getInfoFor(ByVal requestId As String)
        Dim Recordset As DataSet
        Recordset = GetData("sp_AllPackageInfo")
        DisplaySingleData(requestId, Recordset)
    End Sub

    Private Sub DisplaySingleData(ByVal requestId As String, ByVal recordset As DataSet)
        Dim theRow As DataRow
        'theRow = recordset.Tables(0).Rows.Find("Name = '" & requestId & "'")
        'set the checkboxes
    End Sub

    Private Sub LoadDropDown()
        Dim Recordset As DataSet
        Recordset = GetData("sp_AllPackageInfo")
        DisplayData(Recordset)
    End Sub

    Private Sub DisplayData(ByVal Recordset As DataSet)
        Dim aRow As DataRow
        Dim values As ArrayList = New ArrayList()
        Dim aControl As Control

        Me.SelectPackageName.Items.Clear()
        For Each aRow In Recordset.Tables(0).Rows
            Me.SelectPackageName.Items.Add(aRow(1))
        Next aRow

    End Sub

    Private Sub ButtonSave_Click()

    End Sub

    Private Sub ButtonCancel_Click()

    End Sub

End Class
