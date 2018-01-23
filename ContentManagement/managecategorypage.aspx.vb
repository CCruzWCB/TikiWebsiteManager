Imports Management
Imports BPS_BL.BPS


Partial Class ContentManagement_managecategorypage
    Inherits System.Web.UI.Page

    'Load datacubes from database. 
    Private Sub LoadCategoryDataCubes(ByVal CategoryID As Integer)

        Try
            If CategoryID > 0 Then
                Me.ClientDataCube1.Visible = True
                'Me.ClientDataCube2.Visible = True
                'Me.ClientDataCube3.Visible = True
                'Me.ClientDataCube4.Visible = True

                Me.ClientDataCube1.EditMode = True
                Me.ClientDataCube1.LoadDataCube("DataCubeCat" & CategoryID & "_1")
                'Me.ClientDataCube2.EditMode = True
                'Me.ClientDataCube2.LoadDataCube("DataCubeCat" & CategoryID & "_2")
                'Me.ClientDataCube3.EditMode = True
                'Me.ClientDataCube3.LoadDataCube("DataCubeCat" & CategoryID & "_3")
                'Me.ClientDataCube4.EditMode = True
                'Me.ClientDataCube4.LoadDataCube("DataCubeCat" & CategoryID & "_4")

            Else
                Me.ClientDataCube1.Visible = False
                'Me.ClientDataCube2.Visible = False
                'Me.ClientDataCube3.Visible = False
                'Me.ClientDataCube4.Visible = False
            End If
        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error Loading DataCube information for selected Category" & ex.Message.ToString, Management.MessageType.ErrorMessage)


        End Try
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.PageHeader = "Manage 1st level Category Page Content"
            Me.Master.PageTitle = "Manager 1st level Category Content"
            'Me.Master.SetDisplayMessage("Click on one of the data cube below to update content", Management.MessageType.GeneralMessage)

            Me.Master.SetDisplayMessage("Select a Category from the dropdown below to view it's Data Cube elements", Management.MessageType.GeneralMessage)
            Try
                ''Me.ddCategory.DataBind()
                ''Me.ddCategory.Items.Add(Add.ItemtoDropDownList("Select a Category", "0", True))

                Me.ClientDataCube1.Visible = False
                'Me.ClientDataCube2.Visible = False
                'Me.ClientDataCube3.Visible = False

                If Not IsNothing(Request.QueryString("CategoryID")) Then
                    If IsNumeric(Request.QueryString("CategoryID").ToString) Then

                        LoadCategoryDataCubes(Request.QueryString("CategoryID"))
                    End If
                End If

            Catch ex As Exception
                Me.Master.SetDisplayMessage("Error Loading categories: " & ex.Message.ToString, MessageType.ErrorMessage)
                Me.ClientDataCube1.Visible = False
                'Me.ClientDataCube2.Visible = False
                'Me.ClientDataCube3.Visible = False

            End Try
        End If
    End Sub

  
End Class


