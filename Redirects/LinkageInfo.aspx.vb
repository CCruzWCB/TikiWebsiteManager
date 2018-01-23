
Partial Class Redirects_LinkageInfo
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.PageHeader = "Direct Linkage User Guide"
            Me.Master.PageTitle = "Direct Linkage User Guide"
            Me.Master.SetDisplayMessage("Follow the guidelines below to ensure tracking and proper linkage redirection", Management.MessageType.GeneralMessage)

        End If
    End Sub
End Class
