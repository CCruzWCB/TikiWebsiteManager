
Partial Class Security_PasswordSupport
    Inherits System.Web.UI.Page


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Master.SetDisplayMessage("Forgot your password? Follow the direction below to have your password reset and emailed to you. ", Management.MessageType.GeneralMessage)
    End Sub
End Class
