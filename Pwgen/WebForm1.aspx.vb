Public Class WebForm1
    Inherits System.Web.UI.Page
    Protected WithEvents Label1 As System.Web.UI.WebControls.Label
    Protected WithEvents Form1 As System.Web.UI.HtmlControls.HtmlForm
    Protected WithEvents Label2 As System.Web.UI.WebControls.Label
    Protected WithEvents TextBox1 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Label3 As System.Web.UI.WebControls.Label
    Protected WithEvents TextBox2 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button2 As System.Web.UI.WebControls.Button
    Protected WithEvents TextBox3 As System.Web.UI.WebControls.TextBox
    Protected WithEvents Button1 As System.Web.UI.WebControls.Button

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
        'Put user code to initialize the page here
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim lp As Integer
        Dim lengthloop As Integer
        Dim pw As String
        Dim i As Integer
        Dim NoOfOptions As Integer = 3
        Randomize(Timer)
        For lp = 1 To CInt(TextBox1.Text)
            For lengthloop = 1 To CInt(TextBox2.Text)
                'Do
                i = (Rnd(1) * (NoOfOptions - 1)) + 1
                'Loop Until i > 1
                ' 
                Select Case i
                    Case 1
                        pw = pw & Chr((Rnd(1) * 9) + 48) ' Numeric chars
                    Case 2
                        pw = pw & Chr((Rnd(1) * 25) + 65) ' Uppercase chars
                    Case 3
                        pw = pw & Chr((Rnd(1) * 25) + 97) ' Lowercase chars
                    Case 4
                        pw = pw & Chr((Rnd(1) * 15) + 32) ' Other characters (1)
                    Case 5
                        pw = pw & Chr((Rnd(1) * 5) + 91) ' Other characters (2)
                    Case 6
                        pw = pw & Chr((Rnd(1) * 3) + 123) ' Other characters (3)
                End Select
                'pw = pw & " [" & i & "] "
            Next
            TextBox3.Text = TextBox3.Text & vbCrLf & pw
            pw = ""
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TextBox3.Text = ""
    End Sub

End Class
