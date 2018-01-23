Imports System.Data
Imports Management
Imports System.Data.SqlClient

Partial Class MasterPages_Home_Layout1
    Inherits System.Web.UI.MasterPage


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then

        End If

        ' Me.txtSearch.Attributes.Add("onclick", "this.value = '';")

    End Sub

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
    Protected Sub mnuMainCategory_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuMainCategory.DataBound

        Dim mnuItem As MenuItem
        Dim subMenuItem As MenuItem = Nothing



        Try


            Dim iMainCat As Integer
            '   Dim iSubCat As Integer
            If SiteMap.CurrentNode IsNot Nothing Then


                For iMainCat = (mnuMainCategory.Items.Count - 1) To 0 Step -1 ' Looping Backwards



                    If SiteMap.CurrentNode.Title = mnuMainCategory.Items(iMainCat).Text Then
                        mnuMainCategory.Items(iMainCat).Selected = True
                        Exit For
                    End If

                    If SiteMap.CurrentNode.IsDescendantOf(SiteMap.RootNode.ChildNodes(iMainCat)) Then
                        If SiteMap.RootNode.ChildNodes(iMainCat).Title = mnuMainCategory.Items(iMainCat).Text Then
                            mnuMainCategory.Items(iMainCat).Selected = True
                            SiteMapDataSource2.StartFromCurrentNode = False
                            SiteMapDataSource2.StartingNodeOffset = "1"
                            Exit For
                        End If


                    End If



                Next

            End If



        Catch ex As Exception
        Finally
            mnuItem = Nothing
        End Try

    End Sub


    Protected Sub subMenu_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles subMenu.DataBound
        Try


            Dim iMainCat As Integer
            '   Dim iSubCat As Integer
            If SiteMap.CurrentNode IsNot Nothing Then


                For iMainCat = (SubMenu.Items.Count - 1) To 0 Step -1 ' Looping Backwards



                    If SiteMap.CurrentNode.Title = SubMenu.Items(iMainCat).Text Then
                        SubMenu.Items(iMainCat).Selected = True
                        Exit For
                    End If

                    'Check to see which level we are at ...if on a product page...select the product's category
                    If Not IsNothing(SiteMap.CurrentNode.ParentNode.ParentNode) Then
                        'Loop back through and find which the part. 
                        If SiteMap.CurrentNode.ParentNode.Title = SubMenu.Items(iMainCat).Text Then
                            SubMenu.Items(iMainCat).Selected = True
                            SiteMapDataSource2.StartFromCurrentNode = False
                            SiteMapDataSource2.StartingNodeOffset = "1"
                            Exit For
                        End If

                    End If




                Next

            End If

          

        Catch ex As Exception
        Finally
            '            mnuItem = Nothing
        End Try

    End Sub



    Protected Sub subMenu_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles subMenu.PreRender
        If Me.subMenu.Items.Count > 0 Then
            lblMessage.Visible = False
        Else
            lblMessage.Visible = True
        End If

        Try

        Catch ex As Exception

        End Try



    End Sub


End Class

