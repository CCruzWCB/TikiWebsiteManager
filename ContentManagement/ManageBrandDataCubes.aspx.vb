Imports Management
Imports BPS_BL.BPS

Partial Class ContentManagement_ManageBrandDataCubes
    Inherits System.Web.UI.Page

    'Load datacubes from database. 
    'Private Sub LoadSeriesDataCubes(ByVal ProductSeriesID As Integer)

    '    Try
    '        If ProductSeriesID > 0 Then


    '            Me.ClientDataCube1.Visible = True
    '            Me.ClientDataCube2.Visible = True
    '            Me.ClientDataCube3.Visible = True
    '            Me.ClientDataCube4.Visible = True
    '            Me.ClientDataCube5.Visible = True
    '            Me.ClientDataCube6.Visible = True
    '            Me.ClientDataCube7.Visible = True
    '            Me.ClientDataCube8.Visible = True

    '            Me.ClientDataCube1.EditMode = True
    '            Me.ClientDataCube1.LoadDataCube("DataCubeSeries" & ProductSeriesID & "_1")
    '            Me.ClientDataCube2.EditMode = True
    '            Me.ClientDataCube2.LoadDataCube("DataCubeSeries" & ProductSeriesID & "_2")
    '            Me.ClientDataCube3.EditMode = True
    '            Me.ClientDataCube3.LoadDataCube("DataCubeSeries" & ProductSeriesID & "_3")
    '            Me.ClientDataCube4.EditMode = True
    '            Me.ClientDataCube4.LoadDataCube("DataCubeSeries" & ProductSeriesID & "_4")

    '            Me.ClientDataCube5.EditMode = True
    '            Me.ClientDataCube5.LoadDataCube("DataCubeSeries" & ProductSeriesID & "_5")
    '            Me.ClientDataCube6.EditMode = True
    '            Me.ClientDataCube6.LoadDataCube("DataCubeSeries" & ProductSeriesID & "_6")
    '            Me.ClientDataCube7.EditMode = True
    '            Me.ClientDataCube7.LoadDataCube("DataCubeSeries" & ProductSeriesID & "_7")
    '            Me.ClientDataCube8.EditMode = True
    '            Me.ClientDataCube8.LoadDataCube("DataCubeSeries" & ProductSeriesID & "_8")


    '        Else
    '            Me.ClientDataCube1.Visible = False
    '            Me.ClientDataCube2.Visible = False
    '            Me.ClientDataCube3.Visible = False
    '            Me.ClientDataCube4.Visible = False
    '            Me.ClientDataCube5.Visible = False
    '            Me.ClientDataCube6.Visible = False
    '            Me.ClientDataCube7.Visible = False
    '            Me.ClientDataCube8.Visible = False

    '        End If
    '    Catch ex As Exception
    '        Me.Master.SetDisplayMessage("Error Loading DataCube information for selected Series.  " & ex.Message.ToString, Management.MessageType.ErrorMessage)


    '    End Try
    'End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            
            Dim Categoryname As String = ""


            Select Case Request.QueryString("CategoryID").ToString
                Case ConfigurationManager.AppSettings("LamplightBrandCategoryID")
                    Categoryname = "lamplight"
                Case ConfigurationManager.AppSettings("AromaGlowBrandCategoryID")
                    Categoryname = "aromaGlow"
                Case ConfigurationManager.AppSettings("TikiBrandCategoryID")
                    Categoryname = "tiki"
                Case Else
                    'Page.Theme = "Default"
            End Select

            Me.Master.PageHeader = "Manage " & Categoryname & " Category Page Content"
            Me.Master.PageTitle = "Manage " & Categoryname & " Category Content"

            Me.Master.SetDisplayMessage("Click on one of the data cubes below to update content", Management.MessageType.GeneralMessage)
            '            Me.Master.SetDisplayMessage("Select a Series from the dropdown below to view it's Data Cube elements", Management.MessageType.GeneralMessage)


            Try
                ''Me.ddCategory.DataBind()
                ''Me.ddCategory.Items.Add(Add.ItemtoDropDownList("Select a Category", "0", True))

                If Not Page.IsPostBack Then
                    DataCube1.LoadDataCube("DataCube" & Request.QueryString("categoryid") & "_1")
                    DataCube2.LoadDataCube("DataCube" & Request.QueryString("categoryid") & "_2")
                    DataCube3.LoadDataCube("DataCube" & Request.QueryString("categoryid") & "_3")
                    DataCube4.LoadDataCube("DataCube" & Request.QueryString("categoryid") & "_4")
                    DataCube5.LoadDataCube("DataCube" & Request.QueryString("categoryid") & "_5")
                    ' DataCube6.LoadDataCube("DataCube" & Request.QueryString("categoryid") & "_6")
                End If



            Catch ex As Exception
                Me.Master.SetDisplayMessage("Error Loading categories: " & ex.Message.ToString, MessageType.ErrorMessage)
                'Me.ClientDataCube1.Visible = False
                'Me.ClientDataCube2.Visible = False
                'Me.ClientDataCube3.Visible = False

            End Try
        End If
    End Sub


End Class
