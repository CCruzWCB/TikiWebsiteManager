
Partial Class ContentManagement_ManageProductSequence
    Inherits System.Web.UI.Page

    Protected Sub ddProductSeries_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddProductSeries.SelectedIndexChanged
        If Me.ddProductSeries.SelectedValue <> "0" Then
            ActivateSortingPanel(ddProductSeries.SelectedItem.Text, ddProductSeries.SelectedValue)
        End If
    End Sub

    Private Sub ActivateSortingPanel(ByVal selectedItem As String, ByVal SortOrder As String)
        'first need to determine what we are allowing them to sort
        'FORMAT 0.1.0.67.0.430 (ParentSequence.ParentCategoryID.CategorySequence.ProductCategoryID.SeriesSequence.SeriesID)
        Dim aSortingControl As Array = SortOrder.Split(".")

        Me.pSubCategories.Visible = False
        Me.pProducts.Visible = False
         
        If aSortingControl(3) > 0 Then 'ProductSeries was passing...allow sorting of products
            pProducts.Visible = True
            Me.dsGetProductListingBySequence.SelectParameters("ProductCategoryID").DefaultValue = aSortingControl(3)
        Else
            pProducts.Visible = True
            Me.dsGetProductListingBySequence.SelectParameters("ProductCategoryID").DefaultValue = aSortingControl(1)

        End If

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.SetDisplayMessage("Sort a Series of Products below", MessageType.GeneralMessage)
            Me.Master.PageHeader = "Manage Product Sequence"
            Me.Master.PageTitle = "Manage Product Sequence"
        End If
    End Sub
End Class
