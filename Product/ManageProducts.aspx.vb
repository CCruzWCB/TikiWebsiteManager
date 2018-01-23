
Imports BPS_BL.BPS
Partial Class Manager_ManageProducts
    Inherits System.Web.UI.Page


    Protected Sub btnProductSearch_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnProductSearch.Click
        Dim Product As New Product
        Dim sProduct As sProduct = Nothing

        If Me.txtModelNumber.Text <> "" Then
            If Product.GetProduct(sProduct, Me.txtModelNumber.Text) Then
                'Do stuff here
                Response.Redirect(String.Format("ManageProduct.aspx?ProductID={0}", sProduct.ProductID))
            End If
        End If
    End Sub


    'Protected Sub GetProductResults(ByVal ProductID As Long) Handles VisualModelSelector1.ProductSet
    '    Response.Redirect(String.Format("ManageProduct.aspx?ProductID={0}", ProductID))
    'End Sub


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub

   
End Class
