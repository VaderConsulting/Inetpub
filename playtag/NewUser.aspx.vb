
Partial Class NewUser
    Inherits System.Web.UI.Page

    Protected Sub CreateUserWizard1_ActiveStepChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles CreateUserWizard1.ActiveStepChanged
        Dim Status As Integer = CreateUserWizard1.ActiveStepIndex

    End Sub
End Class
