Imports System.Data
Imports Management
Imports System.Data.SqlClient

Partial Class MasterPages_Home_Layout
    Inherits System.Web.UI.MasterPage


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        If Not Page.IsPostBack Then

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


                For iMainCat = (subMenu.Items.Count - 1) To 0 Step -1 ' Looping Backwards



                    If SiteMap.CurrentNode.Title = subMenu.Items(iMainCat).Text Then
                        subMenu.Items(iMainCat).Selected = True
                        Exit For
                    End If

                    'Check to see which level we are at ...if on a product page...select the product's category
                    If Not IsNothing(SiteMap.CurrentNode.ParentNode.ParentNode) Then
                        'Loop back through and find which the part. 
                        If SiteMap.CurrentNode.ParentNode.Title = subMenu.Items(iMainCat).Text Then
                            subMenu.Items(iMainCat).Selected = True
                            SiteMapDataSource2.StartFromCurrentNode = False
                            SiteMapDataSource2.StartingNodeOffset = "1"
                            Exit For
                        End If

                    End If




                Next

            End If

            ''Add Static Menu Items
            'mnuItem = New MenuItem
            'mnuItem.Text = "About us"
            'mnuItem.Selectable = True
            'mnuItem.NavigateUrl = "~/Consumer/AboutUs.aspx"
            'mnuMainCategory.Items.Add(mnuItem)


            'mnuItem = New MenuItem
            'mnuItem.Text = "Buy Online"
            'mnuItem.Selectable = True
            'mnuItem.NavigateUrl = "~/Consumer/BuyOnline.aspx"
            'mnuMainCategory.Items.Add(mnuItem)



            'mnuItem = New MenuItem
            'mnuItem.Text = "Lamplight"
            'mnuItem.Selectable = True
            'mnuItem.NavigateUrl = "~/Consumer/CategoryListings.aspx?Category=7"
            'mnuMainCategory.Items.Add(mnuItem)

            'mnuItem = New MenuItem
            'mnuItem.Text = "Aroma Glow"
            'mnuItem.Selectable = True
            'mnuItem.NavigateUrl = "~/Consumer/CategoryListings.aspx?Category=2"
            'mnuMainCategory.Items.Add(mnuItem)


            ''Load the Recipe categories here

            'LoadMenuRecipe(mnuItem)


            'Dim mnuRecipeAll = New MenuItem

            'mnuRecipeAll.Text = "Browse All"
            'mnuRecipeAll.Value = 0
            'mnuRecipeAll.NavigateUrl = "~\Consumer\RecipeCategory.aspx"
            'mnuItem.ChildItems.Add(mnuRecipeAll)
            'mnuRecipeAll = Nothing

            'Dim mnuRecipeSearch = New MenuItem

            'mnuRecipeSearch.Text = "Search Recipes"
            'mnuRecipeSearch.Value = 0
            'mnuRecipeSearch.NavigateUrl = "~\Consumer\RecipeSearch.aspx"
            'mnuItem.ChildItems.Add(mnuRecipeSearch)
            'mnuRecipeSearch = Nothing



            'Load the E-Commerce categories here

            '     LoadEcommerceCategories(mnuItem)

            'Me.mnuMainCategory.DynamicItemFormatString = Left("{0}" & String.Empty.PadRight(150), 150)

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

