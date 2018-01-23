
Partial Class Redirects_Default
    Inherits System.Web.UI.Page


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.SetDisplayMessage("Welcome to Website Redirect/Linkage Management. Please select one of the options below", Management.MessageType.GeneralMessage)
            Me.Master.PageHeader = "Manage Website Redirect/Linkage Management"
            Me.Master.PageTitle = "Manage Website Redirect/Linkage Management"

        End If
        
    End Sub
End Class
