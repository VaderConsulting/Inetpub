Public Class packaging_rbottom
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
        Dim Table As DataTable, Row As DataRow, Column As DataColumn

        Mode = Page.Request.QueryString("Mode")

        Select Case Mode
            Case "all"
                displayAll()
            Case "add"
                addPackage()
            Case "delete"
                deletePackage()
            Case "status"
                editStatus()
            Case "times"
                editTimes()
        End Select

    End Sub

    Private Sub displayAll()
        Dim Recordset As DataSet
        Recordset = GetData("sp_AllPackageInfo")
        DisplayData(Recordset, 1)
    End Sub

    Private Sub DisplayData(ByVal Recordset As DataSet, ByVal NoOfTableColumns As Int32)
        Dim Table As DataTable, Row As DataRow
        Dim TableColNumber As Int32, Classname As String
        Dim ID As Int32, Name As String, Product As String, Publisher As String, Version As String
        Dim CName As String, SName As String, ClientName As String, Status As String
        Dim States(8) As Boolean, Description As String, SubmitDateTime As String, Media As String, SubmitStatus As String

        ' For each table in the DataSet, output the row values.
        Table = Recordset.Tables(0)
        Response.Write("<table border=""0"" align=""center"" cellspacing=""0"" cellpadding=""0"" cellborder=""0"">" & ControlChars.CrLf)
        Response.Write("    <tr>" & ControlChars.CrLf)

        Response.Write("        <td>&nbsp;</td>" & ControlChars.CrLf)
        Response.Write("        <td class=""BOLD"" width=""250"" height=""20"">Name</td>" & ControlChars.CrLf)
        Response.Write("        <td class=""BOLD"" width=""150"" height=""20"">Packager</td>" & ControlChars.CrLf)
        Response.Write("        <td class=""BOLD"" width=""100"" height=""20"">Client</td>" & ControlChars.CrLf)
        Response.Write("        <td class=""BOLD"" width=""150"" height=""20"">Status</td>" & ControlChars.CrLf)
        Response.Write("        <td class=""BOLD"" width=""80"" height=""20"">Submitted</td>" & ControlChars.CrLf)
        Response.Write("        <td class=""BOLD"" width=""60"" height=""20"">How</td>" & ControlChars.CrLf)
        Response.Write("        <td class=""BOLD"" width=""100"" height=""20"">Result</td>" & ControlChars.CrLf)

        Response.Write("    </tr>" & ControlChars.CrLf)
        Response.Write("    <tr>" & ControlChars.CrLf)

        For Each Row In Table.Rows
            Status = ""
            TableColNumber = TableColNumber + 1

            ID = Row.Item(0)
            Name = Row.Item(1) & ""
            Product = Row.Item(2) & ""
            Publisher = Row.Item(3) & ""
            Version = Row.Item(4) & ""
            CName = Row.Item(5) & ""
            SName = Row.Item(6) & ""
            ClientName = Row.Item(7) & ""
            States(0) = Row.Item(8)  ' WaitSW
            States(1) = Row.Item(9)  ' WaitHW
            States(2) = Row.Item(10) ' WaitRequest
            States(3) = Row.Item(11) ' WaitingPersonnel
            States(4) = Row.Item(12) ' isBuilding
            States(5) = Row.Item(13) ' isTesting
            States(6) = Row.Item(14) ' isValidating
            States(7) = Row.Item(15) ' isComplete
            States(8) = Row.Item(16) ' isPublished
            Description = Row.Item(17) & ""
            SubmitDateTime = Row.Item(18) & ""
            Media = Row.Item(19) & ""
            SubmitStatus = Row.Item(20) & ""

            If States(0) Or States(1) Or States(2) Then
                Status = "Waiting for client"
            End If

            If States(3) Then
                If Status <> "" Then Status = Status & ", "
                Status = Status & "Waiting for staff"
            End If

            If States(4) Then
                If Status <> "" Then Status = Status & ", "
                Status = Status & "Building"
            End If

            If States(5) Then
                If Status <> "" Then Status = Status & ", "
                Status = Status & "Testing"
            End If

            If States(6) Then
                If Status <> "" Then Status = Status & ", "
                Status = Status & "Validating"
            End If

            If States(7) Then
                If Status <> "" Then Status = Status & ", "
                Status = Status & "Published"
            End If

            If States(8) Then
                Select Case Status
                    Case "Published"
                        Status = "Complete and " & Status
                    Case Else
                        Status = "Complete, though " & Status
                End Select
            End If

            If SubmitDateTime <> "" And Not States(4) Then
                SubmitDateTime = FormatDateTime(SubmitDateTime, DateFormat.ShortDate)
            Else
                SubmitDateTime = "N/A"
                Media = "&nbsp;"
                SubmitStatus = "&nbsp;"
            End If

            Select Case SubmitStatus
                Case "Rejected"
                    Classname = "REDTEXT"
                Case "Accepted"
                    Classname = "GREENTEXT"
                Case "Submitted"
                    Classname = "GREYTEXT"
                Case Else
                    Classname = "BLUETEXT"
            End Select

            Response.Write("        <td nowrap valign=""middle"">" & ControlChars.CrLf)
            Response.Write("          <a href=""#""><img src=""images/delete.gif"" alt=""Delete"" border=""0"" width=""18""></a>" & ControlChars.CrLf)
            Response.Write("          <a href=""package_status.aspx?ID=" & ID & """><img src=""images/properties.gif"" alt=""Properties"" border=""0"" width=""18""></a>" & ControlChars.CrLf)
            Response.Write("        </td>" & ControlChars.CrLf)

            Response.Write("        <td class=""" & Classname & """ nowrap width=""250"" height=""20"">" & Name & "</td>" & ControlChars.CrLf)
            Response.Write("        <td class=""" & Classname & """ nowrap width=""150"" height=""20"">" & CName & " " & SName & "</td>" & ControlChars.CrLf)
            Response.Write("        <td class=""" & Classname & """ nowrap width=""100"" height=""20"">" & ClientName & "</td>" & ControlChars.CrLf)
            Response.Write("        <td class=""" & Classname & """ nowrap width=""180"" height=""20"">" & Status & "</td>" & ControlChars.CrLf)
            Response.Write("        <td class=""" & Classname & """ nowrap width=""80"" height=""20"">" & SubmitDateTime & "</td>" & ControlChars.CrLf)
            Response.Write("        <td class=""" & Classname & """ nowrap width=""60"" height=""20"">" & Media & "</td>" & ControlChars.CrLf)
            Response.Write("        <td class=""" & Classname & """ nowrap width=""100"" height=""20"">" & SubmitStatus & "</td>" & ControlChars.CrLf)

            If TableColNumber = NoOfTableColumns Then
                TableColNumber = 0
                Response.Write("    </tr>" & ControlChars.CrLf)
                Response.Write("    <tr>" & ControlChars.CrLf)
            End If
        Next Row
        Response.Write("    </tr>" & ControlChars.CrLf)
        Response.Write("</table>" & ControlChars.CrLf)
    End Sub
    Private Sub editStatus()
        Response.Write("TODO: Status")
    End Sub
    Private Sub editTimes()
        Response.Write("ToDo:  Times")
    End Sub
    Private Sub addPackage()
        Response.Write("ToDo:  Add a Package")
    End Sub
    Private Sub deletePackage()
        Response.Write("ToDo:  Delete a package")
    End Sub
End Class

