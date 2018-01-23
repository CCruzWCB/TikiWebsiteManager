
Partial Class ContentManagement_contentmanagement
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.SetDisplayMessage("Welcome to Page Management. Please select the page you would like to manage below", Management.MessageType.GeneralMessage)
            Me.Master.PageHeader = "Manage Web Site Data Cubes"
            Me.Master.PageTitle = "Manage WEb Site Data Cubes"
        End If
    End Sub
End Class
