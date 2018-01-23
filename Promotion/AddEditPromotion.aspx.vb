
Partial Class Promotion_AddPromotion
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.PageTitle = "Add / Edit Promotions"
            Me.Master.PageHeader = "Add / Edit Promotions"
            Me.Master.SetDisplayMessage("Add or Update fields below", Management.MessageType.GeneralMessage)


        End If
    End Sub


    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        If IsNumeric(Request.QueryString("PID")) AndAlso Request.QueryString("PID") > 0 Then
            Me.fvPromotion.DefaultMode = FormViewMode.Edit
        End If
    End Sub

    Protected Sub fvPromotion_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewInsertedEventArgs) Handles fvPromotion.ItemInserted
        'Me.Master.SetDisplayMessage("Promotion has been added.", Management.MessageType.ConfirmationMessage)
        Response.Redirect("ManagePromotions.aspx")

    End Sub

    Protected Sub fvPromotion_ItemUpdated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewUpdatedEventArgs) Handles fvPromotion.ItemUpdated
        Me.Master.SetDisplayMessage("Promotion has been updated.", Management.MessageType.ConfirmationMessage)
        Response.Redirect("ManagePromotions.aspx")

    End Sub
End Class
