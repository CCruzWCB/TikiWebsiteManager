Imports Management

Imports BPS_BL.BPS

Partial Class _Default
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.SetDisplayMessage("Welcome to " & ConfigurationManager.AppSettings("ApplicationName") & " Management Control Panel", Management.MessageType.GeneralMessage)
            Me.Master.PageTitle = ConfigurationManager.AppSettings("ApplicationName") & " Home "
            Me.Master.PageHeader = ConfigurationManager.AppSettings("ApplicationName") & " Control Panel"
            SetupLastUpdatedItems()
            'lnkStagingDirectory.HRef = ConfigurationManager.AppSettings("ImageBulkUploadDirectory")
        End If
    End Sub

    Private Sub SetupLastUpdatedItems()
        GetProductSeriesImageInfo()
    End Sub

    Private Sub GetProductSeriesImageInfo()
        Dim oProductSeries As New ProductSeries
        Dim sProductSEries As sProductSeries = Nothing
        Try

            If ViewState("SelectedProductSeriesID") > 0 Then
                sProductSEries.ProductSeriesID = ViewState("SelectedProductSeriesID")
                If oProductSeries.GetProductSeries(sProductSEries) Then
                    Me.lblLastProductImageManaged.Text = "Last Selected Product: <a href='Product/manageproductimages.aspx'>" & sProductSEries.Name & "</a>"
                End If
            End If

        Catch ex As Exception

        Finally
            oProductSeries = Nothing
        End Try
    End Sub

End Class
