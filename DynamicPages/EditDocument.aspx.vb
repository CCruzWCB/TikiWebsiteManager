Imports Management
Partial Class FooterContent_EditDocument
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Master.PageHeader = "Manage Dynamic Pages"
        Me.Master.PageTitle = "Manage Dynamic Pages"
        Me.Master.SetDisplayMessage("Welcome to Manage Document Management", MessageType.GeneralMessage)

    End Sub


    Protected Sub fvDocument_ItemUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewUpdatedEventArgs) Handles fvDocument.ItemUpdated
        Response.Redirect("ManageDynamicPages.aspx")
    End Sub
End Class
