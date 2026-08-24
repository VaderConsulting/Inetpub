Imports System.Windows.Forms

Partial Class Login
    Inherits System.Web.UI.Page

    Protected Sub btnNewUser_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnNewUser.Click
        Response.Redirect("newuser.aspx")
    End Sub

    Protected Sub Login1_LoginError(ByVal sender As Object, ByVal e As System.EventArgs) Handles Login1.LoginError
        Login1.CreateUserUrl = "NewUser.aspx"
        Login1.FailureAction = LoginFailureAction.Refresh
        Login1.PasswordRecoveryUrl = "RecoverPassword.aspx"
    End Sub
End Class
