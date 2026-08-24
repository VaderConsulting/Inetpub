Public Class client_details
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
        Dim RequestVariable1 As String, RequestVariable2 As String
        Dim clientid As Int32

        RequestVariable1 = Page.Request.QueryString("ClientId")
        'RequestVariable2 = Page.Request.QueryString("Delete")

        If RequestVariable1 <> "" Then
            clientid = CInt(RequestVariable1)
        Else
            clientid = 0
        End If

        Response.Write("ClientId = " & clientid & vbCrLf)
        'If RequestVariable2 <> "" Then
        '    Delete = CBool(RequestVariable2)
        'Else
        '    Delete = False
        'End If
    End Sub

End Class
