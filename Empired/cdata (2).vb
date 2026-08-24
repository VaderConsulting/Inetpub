Public Class cData

    Public Enum DataTypes
        StoredProcedure = 4
        Text = 1
        TableDirect = 512
    End Enum

    Public Enum DataError
        NOERROR = 0
        NODSN = 1
        CONNECTION = 2
        FILL = 3
    End Enum

    Public DSN As String
    Public ReturnCode As Int32
    Public ErrorMessage As String
    Public RecordCount As Int64

    Public Function Exec(ByVal QueryOrProcedureName As String, Optional ByVal Query_Type As DataTypes = DataTypes.Text) As System.Data.DataSet
        'Dim dreader As SqlClient.SqlDataReader
        Dim Command As New System.Data.SqlClient.SqlCommand()
        Dim Conn As New System.Data.SqlClient.SqlConnection()
        Dim Adapter As New Data.SqlClient.SqlDataAdapter()
        Dim Recordset As New System.Data.DataSet()

        If Trim(DSN) = "" Then
            ReturnCode = DataError.NODSN
            ErrorMessage = "No DSN specified."
            Exit Function
        End If

        Conn.ConnectionString = DSN

        Adapter.SelectCommand = New SqlClient.SqlCommand(QueryOrProcedureName)
        Adapter.TableMappings.Add("Table", "MyTable")

        Try
            Conn.Open()
        Catch
            ReturnCode = DataError.CONNECTION
            ErrorMessage = "Error connecting to Database."
            Exit Function
        End Try

        Command.CommandText = QueryOrProcedureName.ToString
        Command.CommandType = Query_Type
        Command.Connection = Conn

        Adapter.SelectCommand = Command
        ' The connection is now open

        Try
            ' Fill the dataset from the command
            Adapter.Fill(Recordset)
        Catch
            ReturnCode = DataError.FILL
            ErrorMessage = "Error filling Data Adapter."
            Exit Function
        End Try

        RecordCount = Recordset.Tables(0).Rows.Count

        ' Close the connection
        Conn.Close()

        ' Set return code
        ReturnCode = DataError.NOERROR
        ErrorMessage = "No Error"
        ' Done
        Return Recordset
    End Function
End Class
