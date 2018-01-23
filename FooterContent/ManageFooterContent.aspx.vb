Imports Management
Partial Class FooterContent_ManageFooterContent
    Inherits System.Web.UI.Page

    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGo.Click
        Response.Redirect(String.Format("EditDocument.aspx?doc_id={0}", Me.ddlDocument.SelectedValue))
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Master.PageHeader = "Manage Static Footer Content"
        Me.Master.PageTitle = "Manage Static Footer Content"
        Me.Master.SetDisplayMessage("Welcome to Manage Static Footer Content Management", MessageType.GeneralMessage)

    End Sub
End Class
