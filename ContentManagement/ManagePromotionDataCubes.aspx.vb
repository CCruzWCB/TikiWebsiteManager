Imports Management
Imports BPS_BL.BPS

Partial Class ContentManagement_ManagePromotionDataCubes
    Inherits System.Web.UI.Page
    'Load datacubes from database. 
    Private Sub LoadPromotionDataCubes(ByVal PromotionID As Integer)

        Try
            If PromotionID > 0 Then


                Me.ClientDataCube1.Visible = True
               
                Me.ClientDataCube1.EditMode = True
                Me.ClientDataCube1.LoadDataCube("DataCubePromo" & PromotionID & "_1")
               

            Else
                Me.ClientDataCube1.Visible = False
               
            End If
        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error Loading DataCube information for selected Promotion.  " & ex.Message.ToString, Management.MessageType.ErrorMessage)


        End Try
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.PageHeader = "Manage Promotion Page Content"
            Me.Master.PageTitle = "Manager Promotion Content"
            Me.Master.SetDisplayMessage("Click on one of the data cubes below to update content", Management.MessageType.GeneralMessage)

            Me.Master.SetDisplayMessage("Select a Promotion from the dropdown below to view it's Data Cube elements", Management.MessageType.GeneralMessage)
            Try
                ''Me.ddCategory.DataBind()
                ''Me.ddCategory.Items.Add(Add.ItemtoDropDownList("Select a Category", "0", True))

                Me.ClientDataCube1.Visible = False
            

                If Not IsNothing(Request.QueryString("PromotionID")) Then
                    If IsNumeric(Request.QueryString("PromotionID").ToString) Then

                        LoadPromotionDataCubes(Request.QueryString("PromotionID"))
                    End If
                End If

            Catch ex As Exception
                Me.Master.SetDisplayMessage("Error Loading categories: " & ex.Message.ToString, MessageType.ErrorMessage)
                Me.ClientDataCube1.Visible = False
            
            End Try
        End If
    End Sub


End Class
