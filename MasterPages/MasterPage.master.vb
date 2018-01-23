Imports Management

Partial Class MasterPages_MasterPage
    Inherits System.Web.UI.MasterPage

    'Private Variables 
    Private _PageTitle As String
    Private _DisplayMessage As String
    Private _PageHeader As String


    'Public Properties


    Public Property PageHeader() As String
        Get
            Return _PageHeader
        End Get
        Set(ByVal value As String)
            _PageHeader = value
            Me.lblMasterPageHeader.Text = value
        End Set
    End Property

    Public Property PageTitle() As String
        Get
            Return _PageTitle
        End Get
        Set(ByVal value As String)
            _PageTitle = value
            Me.myPageTitle.Text = _PageTitle
        End Set
    End Property

    Public Property DisplayMessage() As String
        Get
            Return _DisplayMessage
        End Get
        Set(ByVal value As String)
            _DisplayMessage = value
        End Set
    End Property


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        ' AddHandler SiteMap.SiteMapResolve, AddressOf Me.ExpandForumPaths

        Dim currentNode As SiteMapNode = SiteMap.CurrentNode


    End Sub

    Protected Sub imgLogo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLogo.Click
        Response.Redirect("~/default.aspx")
    End Sub

    Public Sub SetDisplayMessage(ByVal DisplayMessage As String, ByVal MessageType As MessageType)
        Me.lblDisplayMessage.CssClass = ""
        Me.lblDisplayMessage.Text = ""

        Select Case MessageType
            Case MessageType.ErrorMessage
                Me.lblDisplayMessage.CssClass = "errorbox"
            Case MessageType.GeneralMessage
                Me.lblDisplayMessage.CssClass = "general"
            Case MessageType.SyntaxMessage
                Me.lblDisplayMessage.CssClass = "syntaxerror"
            Case MessageType.instructionMessage
                Me.lblDisplayMessage.CssClass = "instructions"
            Case MessageType.WizardMessage
                Me.lblDisplayMessage.CssClass = "wizard"
            Case MessageType.ConfirmationMessage
                Me.lblDisplayMessage.CssClass = "confirmation"
        End Select

        Me.lblDisplayMessage.Text = DisplayMessage

    End Sub

    Private Function ExpandForumPaths(ByVal sender As Object, ByVal e As SiteMapResolveEventArgs) As SiteMapNode
        
        Try
            'Create instance of current site map Node
            Dim currentNode As SiteMapNode = SiteMap.CurrentNode.Clone(True)
            Dim tempNode As SiteMapNode = currentNode


            Dim ProductSeriesID As Integer = Me.GetProductSeriesID
            Dim ProductCategoryID As Integer = Me.GetProductCategoryID



            tempNode = tempNode.ParentNode
            If Not (0 = ProductSeriesID) And Not (IsNothing(tempNode)) Then
                tempNode.Url = tempNode.Url & "?ProductSeriesID=" & ProductCategoryID
            End If

            tempNode = tempNode.ParentNode
            If Not (0 = ProductCategoryID) And Not (IsNothing(tempNode)) Then
                tempNode.Url = tempNode.Url & "?ProductCategoryID=" & ProductCategoryID
            End If

            Return currentNode

        Catch ex As Exception

        End Try

    End Function


    Private Function GetProductSeriesID() As Integer
        'Return Profile.SelectedProductSeriesID
        Return ViewState("SelectedProductSeriesID")
    End Function

    Private Function GetProductCategoryID() As Integer
        'Return Profile.SelectedProductCategoryID
        Return ViewState("SelectedProductCategoryID")
    End Function

  
End Class

