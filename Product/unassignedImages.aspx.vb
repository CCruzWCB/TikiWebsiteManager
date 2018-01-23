
Partial Class Product_unassignedImages
    Inherits System.Web.UI.Page

    
    Protected Sub dlMissingProductImages_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataListCommandEventArgs) Handles dlMissingProductImages.ItemCommand
        If e.CommandArgument > 0 Then

            Response.Redirect(String.Format("manageproductimages.aspx?ProductID={0}", e.CommandArgument))
        End If


    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.SetDisplayMessage("There are " & Me.dlMissingProductImages.Items.Count & " unassigned product images", Management.MessageType.GeneralMessage)
            Me.Master.PageHeader = "Unassigned Product Image Management"
            Me.Master.PageTitle = "Unassigned Product Images"
        End If
    End Sub
End Class
