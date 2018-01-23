Imports Management
Partial Class MasterPages_managehomepage
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.SetDisplayMessage("Click on the highlighted area to change content", Management.MessageType.GeneralMessage)
            Me.Master.PageHeader = " Manage Home Page Content "

        End If



    End Sub
End Class
