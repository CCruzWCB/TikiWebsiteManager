Imports System.Web.Security
Imports System.Data
Imports System.Data.SqlClient
Imports System.Web.SiteMapProvider
Imports Management


Partial Class C_Layout
    Inherits System.Web.UI.MasterPage
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
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then

        End If



    End Sub

    Public Sub SetDisplayMessage(ByVal DisplayMessage As String, ByVal MessageType As MessageType)
        Me.Master.SetDisplayMessage(DisplayMessage, MessageType)
    End Sub

    Protected Sub dlCategory_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataListItemEventArgs) Handles dlCategory.ItemDataBound
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            If Not IsNothing(Request.QueryString("cat")) Then
                If IsNumeric(Request.QueryString("cat")) Then
                    If e.Item.DataItem("ProductCategoryID") = Request.QueryString("cat") Then
                        CType(e.Item.FindControl("lnkCategory"), HyperLink).CssClass = "yellowboxon"
                    Else
                        CType(e.Item.FindControl("lnkCategory"), HyperLink).CssClass = "yellowbox"
                    End If
                End If
            End If

        End If
    End Sub
End Class

