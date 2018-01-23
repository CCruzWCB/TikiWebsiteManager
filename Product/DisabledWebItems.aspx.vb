
Partial Class Product_DisabledWebItems
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.PageHeader = "Manage Items with disabled ""Add To Cart"" button"
            Me.Master.PageTitle = "Manage Items with disabled ""Add To Cart"" button"
            Me.Master.SetDisplayMessage("View the list of items where the ""Add to Cart"" buton has been disabled below.  To add item, type number and click add.  Click remove on line item to remove.", MessageType.GeneralMessage)
        End If
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
        dsDisableItems.UpdateParameters("ProductModelNumber").DefaultValue = txtItem.Text.ToString.Trim

        dsDisableItems.Update()

        gvDisabledAddToCartItems.DataBind()

    End Sub

    Protected Sub dsDisableItems_Deleted(sender As Object, e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsDisableItems.Deleted
        Master.SetDisplayMessage("The disable flag has been cleared from this item", MessageType.GeneralMessage)
        gvDisabledAddToCartItems.DataBind()
    End Sub

    Protected Sub dsDisableItems_Updated(sender As Object, e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsDisableItems.Updated

        Master.SetDisplayMessage("The disable flag has be added for this item", MessageType.GeneralMessage)
        txtItem.Text = ""


    End Sub

    Protected Sub gvDisabledAddToCartItems_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvDisabledAddToCartItems.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim delButton As LinkButton = e.Row.Cells(3).Controls(0)

            delButton.OnClientClick = "return confirm('You are about to remove the disable flag for this item. Are you sure??');"
        End If


    End Sub

    Protected Sub gvDisabledAddToCartItems_RowDeleted(sender As Object, e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles gvDisabledAddToCartItems.RowDeleted
        'dsDisableItems.Delete()

    End Sub
End Class
