
Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_LoadComplete(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.LoadComplete
        Response.Redirect("login.aspx")
    End Sub
End Class
