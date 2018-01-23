
Partial Class Promotion_ManagePromotions
    Inherits System.Web.UI.Page
    Public Function ConvertStatus(ByVal bValue As Boolean) As String

        If bValue Then
            Return "Enabled"
        Else
            Return "Disabled"

        End If
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.PageHeader = "Manage Promotions"
            Me.Master.PageTitle = "Manage Promotions"
            Me.Master.SetDisplayMessage("Click on one of the Promotions below to update", Management.MessageType.GeneralMessage)

        End If
    End Sub
End Class
