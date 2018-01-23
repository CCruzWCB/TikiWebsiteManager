Imports Management

Partial Class MasterPages_Series_Layout
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

    Protected Sub ddlSeries_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSeries.SelectedIndexChanged
        Response.Redirect(String.Format("ManageSeriesDataCubes.aspx?ProductSeriesID={0}", ddlSeries.SelectedValue))
    End Sub
End Class





