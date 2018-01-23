Imports Management

Partial Class MasterPages_Promo_Layout
    Inherits System.Web.UI.MasterPage


    Public Sub SetDisplayMessage(ByVal DisplayMessage As String, ByVal MessageType As MessageType)
        Me.Master.SetDisplayMessage(DisplayMessage, MessageType)
    End Sub
    Public Property PageHeader() As String
        Get
            Return Me.Master.PageHeader
        End Get
        Set(ByVal value As String)
            Me.Master.PageHeader = value

        End Set
    End Property

    Public Property PageTitle() As String
        Get
            Return Me.Master.PageTitle
        End Get
        Set(ByVal value As String)
            Me.Master.PageTitle = value

        End Set
    End Property

    Protected Sub ddlPromotions_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlPromotions.SelectedIndexChanged
        Response.Redirect(String.Format("ManagePromotionDataCubes.aspx?PromotionID={0}", ddlPromotions.SelectedValue))
    End Sub
End Class

