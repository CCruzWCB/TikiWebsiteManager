Imports System.Data
Imports Management
Imports System.Data.SqlClient

Partial Class MasterPages_BuyOnline_Layout
    Inherits System.Web.UI.MasterPage


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then
            LoadCategories()
        End If

        ' Me.txtSearch.Attributes.Add("onclick", "this.value = '';")

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
    Public Sub SetDisplayMessage(ByVal DisplayMessage As String, ByVal MessageType As MessageType)
        Me.Master.SetDisplayMessage(DisplayMessage, MessageType)
    End Sub


    Private Sub LoadCategories()
        'For each main categories...append to table. 


    End Sub


    'Protected Sub btnGo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnGo.Click
    '    If Trim(Me.txtSearch.Text) <> "" Then
    '        Response.Redirect(String.Format("SearchResults.aspx?Search={0}", Trim(Me.txtSearch.Text)))
    '    End If
    'End Sub


    'Private Sub SetLoginStatus()
    '    Try

    '        If Page.User.Identity.IsAuthenticated Then
    '            Dim Name As String = ""
    '            If Not IsNothing(Request.Cookies("UserName")) Then
    '                If Request.Cookies("UserName").Value <> "" Then
    '                    Name = Request.Cookies("UserName").Value
    '                    loginstatus1.LogoutText = "logged in as " & Name & "[logout]"
    '                    loginstatus1.ToolTip = "Click to Logout"
    '                End If
    '            Else
    '                'Broswer doesn't support cookies. 
    '                loginstatus1.LogoutText = "[logout]"
    '                loginstatus1.ToolTip = "Click to Logout"
    '            End If


    '        End If

    '    Catch ex As Exception

    '    End Try

    'End Sub

    'Public Function GetCustomerName() As String
    '    If Page.User.Identity.IsAuthenticated Then

    '        If Not IsNothing(Request.Cookies("UserName")) Then
    '            Return Request.Cookies("UserName").Value
    '        Else
    '            Return ""
    '        End If
    '    Else
    '        Return ""
    '    End If
    'End Function

    'Public Sub SetCustomerName(ByVal Name As String)
    '    If Page.User.Identity.IsAuthenticated Then
    '        Response.Cookies("UserName").Value = Name
    '        'TODO Response.Cookies ("UserName").Expires =  'WEBConfig value 
    '    End If
    'End Sub


End Class

