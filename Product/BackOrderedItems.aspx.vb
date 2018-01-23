
Partial Class Product_BackorderedItems
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.PageHeader = "Manage Backordered Items"
            Me.Master.PageTitle = "Manage Backordered Items"
            Me.Master.SetDisplayMessage("View the list of backordered items below.  To add item, type number and click add.  Click remove on line item to remove.", MessageType.GeneralMessage)
        End If
    End Sub

    Protected Sub btnAdd_Click(sender As Object, e As System.EventArgs) Handles btnAdd.Click
        dsBackOrderedItems.UpdateParameters("ProductModelNumber").DefaultValue = txtItem.Text.ToString.Trim

        dsBackOrderedItems.Update()

        gvBackorderItems.DataBind()

    End Sub

    Protected Sub dsBackOrderedItems_Deleted(sender As Object, e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsBackOrderedItems.Deleted
        Master.SetDisplayMessage("Item removed from backorder", MessageType.GeneralMessage)
        gvBackorderItems.DataBind()
    End Sub

    Protected Sub dsBackOrderedItems_Updated(sender As Object, e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsBackOrderedItems.Updated

        Master.SetDisplayMessage("Item added to backorder", MessageType.GeneralMessage)
        txtItem.Text = ""


    End Sub

    Protected Sub gvBackorderItems_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvBackorderItems.RowDataBound

        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim delButton As LinkButton = e.Row.Cells(3).Controls(0)

            delButton.OnClientClick = "return confirm('You are about to remove this item from backorder. Are you sure??');"
        End If

     
    End Sub

    Protected Sub gvBackorderItems_RowDeleted(sender As Object, e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles gvBackorderItems.RowDeleted
        'dsBackOrderedItems.Delete()

    End Sub
End Class
