Imports Management

Partial Class MetaData_ManageMetaData
    Inherits System.Web.UI.Page

    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        Me.Master.SetDisplayMessage("Welcome to Meta Data Management page", MessageType.GeneralMessage)
        Me.Master.PageHeader = "Manage Meta Data"
        Me.Master.PageTitle = "Manage Meta Data"
    End Sub

End Class
