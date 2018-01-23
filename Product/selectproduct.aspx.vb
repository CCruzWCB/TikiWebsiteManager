Imports BPS_BL.BPS
Imports Management



Partial Class Product_selectproduct
    Inherits System.Web.UI.Page

    Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGo.Click
        Dim Product As New Product
        Dim sProduct As sProduct = Nothing

        Try

            If Me.txtProductNumber.Text <> "" Then
                If Product.GetProduct(sProduct, Me.txtProductNumber.Text) Then
                    'Do stuff here
                    Response.Redirect(String.Format("ManageProduct.aspx?ProductID={0}", sProduct.ProductID))
                    'Response.Redirect(String.Format("ManageProductImages.aspx?ProductID={0}", sProduct.ProductID))
                End If
            End If

        Catch ex As Exception
        Finally
            Product = Nothing
        End Try
    End Sub


    Protected Sub ddlProducts_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlProducts.SelectedIndexChanged
        Response.Redirect(String.Format("ManageProduct.aspx?ProductID={0}", Me.ddlProducts.SelectedValue))
        'Response.Redirect(String.Format("ManageProductImages.aspx?ProductID={0}", Me.ddlProducts.SelectedValue))

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.SetDisplayMessage("Select or enter a Product model below", MessageType.GeneralMessage)
            Me.Master.PageHeader = "Select Product"
            Me.Master.PageTitle = "Select Product"
        End If

    End Sub
End Class
