
Partial Class ClientSide_Default
    Inherits System.Web.UI.Page

    Public Sub SetDataCubeID()
        'Select Case Session("CategoryID")
        '    Case GLDepartment.GrillsAndCookers
        '        ClientDataCube1.DataCubeID = "DataCube6"
        '        DataCube2.DataCubeID = "DataCube7"
        '        DataCube3.DataCubeID = "DataCube8"
        '        DataCube4.DataCubeID = "DataCube21"
        '        DataCube5.DataCubeID = "DataCubeGrillHeader"

        '    Case GLDepartment.GrillingAndKitchenTools
        '        DataCube1.DataCubeID = "DataCube9"
        '        DataCube2.DataCubeID = "DataCube10"
        '        DataCube3.DataCubeID = "DataCube11"
        '        DataCube4.DataCubeID = "DataCube22"
        '        DataCube5.DataCubeID = "DataCubeGrillToolsHeader"
        '    Case GLDepartment.SaucesAndRubs
        '        DataCube1.DataCubeID = "DataCube12"
        '        DataCube2.DataCubeID = "DataCube13"
        '        DataCube3.DataCubeID = "DataCube14"
        '        DataCube4.DataCubeID = "DataCube23"
        '        DataCube5.DataCubeID = "DataCubeSaucesHeader"
        '    Case GLDepartment.OutDoorLife
        '        DataCube1.DataCubeID = "DataCube15"
        '        DataCube2.DataCubeID = "DataCube16"
        '        DataCube3.DataCubeID = "DataCube17"
        '        DataCube4.DataCubeID = "DataCube24"
        '        DataCube5.DataCubeID = "DataCubeOutdoorlifelHeader"
        '    Case GLDepartment.GrillParts
        '        DataCube1.DataCubeID = "DataCube18"
        '        DataCube2.DataCubeID = "DataCube19"
        '        DataCube3.DataCubeID = "DataCube20"
        '        DataCube4.DataCubeID = "DataCube24"
        '        DataCube5.DataCubeID = ""
        'End Select
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            'Me.ClientDataCube1.Enabled = False
            'Me.ClientDataCube2.Enabled = False
            'Me.ClientDataCube3.Enabled = False
            'Me.ClientDataCube4.Enabled = False
        End If


    End Sub



    Protected Sub Category_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Category.SelectedIndexChanged
        If Me.Category.SelectedValue > 0 Then
            Me.ClientDataCube1.CategoryID = Me.Category.SelectedValue
            Me.ClientDataCube2.CategoryID = Me.Category.SelectedValue
            Me.ClientDataCube3.CategoryID = Me.Category.SelectedValue
            Me.ClientDataCube4.CategoryID = Me.Category.SelectedValue

        End If
    End Sub

    Private Sub Reset()
        Me.ClientDataCube1.CategoryID = 0
        Me.ClientDataCube1.DataCubeID = 0
        Me.ClientDataCube2.CategoryID = 0
        Me.ClientDataCube2.DataCubeID = 0
        Me.ClientDataCube3.CategoryID = 0
        Me.ClientDataCube3.DataCubeID = 0
        Me.ClientDataCube4.CategoryID = 0
        Me.ClientDataCube4.DataCubeID = 0
    End Sub

    Protected Sub btnReset_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnReset.Click
        Reset()
        'Response.Redirect("default.aspx")
    End Sub
End Class
