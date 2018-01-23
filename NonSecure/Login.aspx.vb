
Partial Class Security_Login
    Inherits System.Web.UI.Page




    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim ApplicationName As String = ConfigurationManager.AppSettings("ApplicationName").ToString
        Me.Master.SetDisplayMessage("Welcome to the " & ApplicationName & " Site Management Tool. Please login in below", Management.MessageType.GeneralMessage)
    End Sub

    Protected Sub btnLogin_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLogin.Click
        If Me.txtPassword.Text = ConfigurationManager.AppSettings("Application_Password").ToString And Me.txtUserName.Text = ConfigurationManager.AppSettings("Application_UserName").ToString Then
            FormsAuthentication.SetAuthCookie(ConfigurationManager.AppSettings("Application_UserName").ToString, Me.chkRememberMe.Checked)
            Response.Redirect("../default.aspx")
        Else
            Me.Master.SetDisplayMessage("Login was incorrect! Please check your username and/or password and try again.", MessageType.ErrorMessage)
        End If
    End Sub
End Class
