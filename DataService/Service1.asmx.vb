Imports System.Web.Services
Imports System.Data
Imports System.Data.SqlClient

<WebService(Namespace:="http://www.empired.com/")> _
Public Class Service1
    Inherits System.Web.Services.WebService

#Region " Web Services Designer Generated Code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Web Services Designer.
        InitializeComponent()

        'Add your own initialization code after the InitializeComponent() call

    End Sub

    'Required by the Web Services Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Web Services Designer
    'It can be modified using the Web Services Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        'CODEGEN: This procedure is required by the Web Services Designer
        'Do not modify it using the code editor.
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

#End Region

    <WebMethod()> Public Function GetData(ByVal ConnString As String, ByVal SelectQuery As String, ByRef Result As String) As String
        Dim Conn As New SqlConnection(ConnString)
        Dim Comm As New SqlCommand(SelectQuery, Conn)

        Result = "Complete"

        Try
            Comm.Connection.Open()
            Dim Reader As SqlDataReader = Comm.ExecuteReader(CommandBehavior.CloseConnection)
            Dim ds As New DataSet("MyDataset")


            Dim da As New SqlDataAdapter(SelectQuery, Conn)
            da.Fill(ds)
            Return ds.Tables.Count.ToString + ConnString + SelectQuery
            Comm.Connection.Close()

        Catch e As Exception
            Result = e.Message
        Finally
        End Try


    End Function

End Class
