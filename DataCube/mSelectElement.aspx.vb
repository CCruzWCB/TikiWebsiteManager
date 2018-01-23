Imports Management
Imports System.Data
Imports System.Web
Imports System.Web.UI

Public Class DataCube_mSelectElement
    Inherits System.Web.UI.Page
    Dim lstFilter1 As New DropDownList
    Dim lstFilter2 As New DropDownList
    Dim sDataSource As sDataSource
    '    Protected WithEvents DsElement1 As dsElement
    Protected WithEvents txtSearch As System.Web.UI.WebControls.TextBox
    Protected WithEvents btnCancel As System.Web.UI.WebControls.Button
    Protected WithEvents btnSave As System.Web.UI.WebControls.Button
    
    Dim oDataCubeObject As New DataCube

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '  Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Me.SqlSelectCommand1 = New System.Data.SqlClient.SqlCommand
        Me.cnnElement = New System.Data.SqlClient.SqlConnection
        Me.SqlInsertCommand1 = New System.Data.SqlClient.SqlCommand
        Me.daElement = New System.Data.SqlClient.SqlDataAdapter
        '
        'SqlSelectCommand1
        '
        Me.SqlSelectCommand1.CommandText = "SELECT Table_View_Name, Table_KeyName1, Table_KeyName2, Table_KeyName3, Table_Dis" & _
        "playName, Table_SelectionField, DataItemType_Name, DataCube_ID, DataCube_Element" & _
        "ID, DataItemType_ID FROM vwDataSource"
        Me.SqlSelectCommand1.Connection = Me.cnnElement
        '
        'cnnElement
        '
        Me.cnnElement.ConnectionString = ConfigurationManager.ConnectionStrings("BaseDBConnection").ConnectionString

        'CType(configurationAppSettings.GetValue("DBConnection", GetType(System.String)), String)
        '
        'SqlInsertCommand1
        '
        Me.SqlInsertCommand1.CommandText = "INSERT INTO vwDataSource(Table_View_Name, Table_KeyName1, Table_KeyName2, Table_K" & _
        "eyName3, Table_DisplayName, Table_SelectionField, DataItemType_Name, DataCube_ID" & _
        ", DataCube_ElementID, DataItemType_ID) VALUES (@Table_View_Name, @Table_KeyName1" & _
        ", @Table_KeyName2, @Table_KeyName3, @Table_DisplayName, @Table_SelectionField, @" & _
        "DataItemType_Name, @DataCube_ID, @DataCube_ElementID, @DataItemType_ID); SELECT " & _
        "Table_View_Name, Table_KeyName1, Table_KeyName2, Table_KeyName3, Table_DisplayNa" & _
        "me, Table_SelectionField, DataItemType_Name, DataCube_ID, DataCube_ElementID, Da" & _
        "taItemType_ID FROM vwDataSource"
        Me.SqlInsertCommand1.Connection = Me.cnnElement
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Table_View_Name", System.Data.SqlDbType.VarChar, 50, "Table_View_Name"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Table_KeyName1", System.Data.SqlDbType.VarChar, 50, "Table_KeyName1"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Table_KeyName2", System.Data.SqlDbType.VarChar, 50, "Table_KeyName2"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Table_KeyName3", System.Data.SqlDbType.VarChar, 50, "Table_KeyName3"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Table_DisplayName", System.Data.SqlDbType.VarChar, 50, "Table_DisplayName"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Table_SelectionField", System.Data.SqlDbType.VarChar, 50, "Table_SelectionField"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DataItemType_Name", System.Data.SqlDbType.VarChar, 50, "DataItemType_Name"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DataCube_ID", System.Data.SqlDbType.VarChar, 50, "DataCube_ID"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DataCube_ElementID", System.Data.SqlDbType.VarChar, 50, "DataCube_ElementID"))
        Me.SqlInsertCommand1.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DataItemType_ID", System.Data.SqlDbType.Int, 4, "DataItemType_ID"))
        '
        'daElement
        '
        Me.daElement.InsertCommand = Me.SqlInsertCommand1
        Me.daElement.SelectCommand = Me.SqlSelectCommand1
        Me.daElement.TableMappings.AddRange(New System.Data.Common.DataTableMapping() {New System.Data.Common.DataTableMapping("Table", "vwDataSource", New System.Data.Common.DataColumnMapping() {New System.Data.Common.DataColumnMapping("Table_View_Name", "Table_View_Name"), New System.Data.Common.DataColumnMapping("Table_KeyName1", "Table_KeyName1"), New System.Data.Common.DataColumnMapping("Table_KeyName2", "Table_KeyName2"), New System.Data.Common.DataColumnMapping("Table_KeyName3", "Table_KeyName3"), New System.Data.Common.DataColumnMapping("Table_DisplayName", "Table_DisplayName"), New System.Data.Common.DataColumnMapping("Table_SelectionField", "Table_SelectionField"), New System.Data.Common.DataColumnMapping("DataItemType_Name", "DataItemType_Name"), New System.Data.Common.DataColumnMapping("DataCube_ID", "DataCube_ID"), New System.Data.Common.DataColumnMapping("DataCube_ElementID", "DataCube_ElementID"), New System.Data.Common.DataColumnMapping("DataItemType_ID", "DataItemType_ID")})})
        '
        

    End Sub

    Protected WithEvents SqlSelectCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents SqlInsertCommand1 As System.Data.SqlClient.SqlCommand
    Protected WithEvents cnnElement As System.Data.SqlClient.SqlConnection
    Protected WithEvents daElement As System.Data.SqlClient.SqlDataAdapter

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here


        If Page.IsPostBack Then

        Else

            Dim strDataCubeID As String = Request.QueryString("did")
            Dim strTypeID As Integer = CType(Request.QueryString("tid"), Integer)
            Me.hdnControlName.Value = Request.QueryString("ctr").ToString

            sDataSource = oDataCubeObject.GetDataSource(strDataCubeID, strTypeID)
            ViewState("FieldValue") = sDataSource.TableValueField
            ViewState("TableDisplayField") = sDataSource.TableDisplayField
            ViewState("TableViewName") = sDataSource.TableViewName

            BuildFilter(strTypeID, sDataSource.tableKey2, sDataSource.tableKey2, sDataSource.tableKey2, sDataSource.tableKey2)
        End If



    End Sub


    Private Sub BuildFilter(ByVal DataCubeType As Integer, ByVal Field1 As String, ByVal Field1Value As String, ByVal Field2 As String, ByVal Field2Value As String)
        Dim strSQL As String

        'strSQL = "select Left(" & sDataSource.TableDisplayField & ",20) +'(' + " & sDataSource.TableValueField & " + ')' as [DisplayField] , " & sDataSource.TableValueField & " From " & sDataSource.TableViewName & " order by " & sDataSource.TableValueField
        strSQL = "select distinct convert(varchar," & sDataSource.tableValueField & " ) + ' - ('+ Left(" & sDataSource.tableDisplayField & ",50) + '...)' as [DisplayField] , " & sDataSource.tableValueField & "  From " & sDataSource.tableViewName & " order by " & sDataSource.tableValueField

        Dim oCmd As New SqlClient.SqlCommand
        Dim oDataSet As New DataSet

        'Determine which connection string to use
        Dim ConnectionString As String = ""

        'Check to see which database should be checked. 
        Select Case DataCubeType
            'Check BPS Database for these items 
            Case Management.DataCubeType.Category, Management.DataCubeType.Product, Management.DataCubeType.ProductSeries, Management.DataCubeType.VirtualCategory
                ConnectionString = ConfigurationManager.ConnectionStrings("BaseDBConnection").ConnectionString
                'Check CB web site Database for these items 
            Case Management.DataCubeType.Promotion
                ConnectionString = ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ConnectionString
            Case Else
                ConnectionString = ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ConnectionString
        End Select

        Me.cnnElement.ConnectionString = ConnectionString

        oCmd.CommandText = strSQL
        oCmd.CommandType = CommandType.Text
        oCmd.Connection = Me.cnnElement
        Me.cnnElement.Open()


        Me.daElement.SelectCommand = oCmd
        Me.daElement.Fill(oDataSet)
        'Dim oTB As New TrickBag.Convert
        'Dim oReader As SqlDataReader
        'oReader = oCmd.ExecuteReader


        'oDataSet = oTB.DataReaderToDataSet(oReader)
        'oTB = Nothing
        Me.lstElement.DataSource = oDataSet
        Me.lstElement.DataMember = oDataSet.Tables(0).TableName


        Me.lstElement.DataTextField = oDataSet.Tables(0).Columns(0).ColumnName ' "DisplayField"
        Me.lstElement.DataValueField = oDataSet.Tables(0).Columns(1).ColumnName  'item_id
        Me.lstElement.DataBind()



    End Sub


    'Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim strSQL As String

    '    'strSQL = "select Left(" & sDataSource.TableDisplayField & ",20) +'(' + " & sDataSource.TableValueField & " + ')' as [DisplayField] , " & sDataSource.TableValueField & " From " & sDataSource.TableViewName & " order by " & sDataSource.TableValueField
    '    strSQL = "select distinct convert(varchar," & ViewState("FieldValue") & " ) + '('+ Left(" & ViewState("TableDisplayField") & ",20) + '...)' as [DisplayField] , " & ViewState("FieldValue") & "  From " & ViewState("TableViewName") & _
    '    " Where " & ViewState("TableDisplayField") & " Like   '%" & Me.txtSearch.Text & "%'" ' order by " & ViewState("TableDisplayField")

    '    Dim oCmd As New SqlClient.SqlCommand
    '    Dim oDataSet As New DataSet

    '    oCmd.CommandText = strSQL
    '    oCmd.CommandType = CommandType.Text
    '    oCmd.Connection = Me.cnnElement
    '    Me.cnnElement.Open()


    '    Me.daElement.SelectCommand = oCmd
    '    Me.daElement.Fill(oDataSet)
    '    'Dim oTB As New TrickBag.Convert
    '    'Dim oReader As SqlDataReader
    '    'oReader = oCmd.ExecuteReader


    '    'oDataSet = oTB.DataReaderToDataSet(oReader)
    '    'oTB = Nothing
    '    Me.lstElement.DataSource = oDataSet
    '    Me.lstElement.DataMember = oDataSet.Tables(0).TableName


    '    Me.lstElement.DataTextField = oDataSet.Tables(0).Columns(0).ColumnName ' "DisplayField"
    '    Me.lstElement.DataValueField = oDataSet.Tables(0).Columns(1).ColumnName  'item_id
    '    Me.lstElement.DataBind()


    'End Sub
End Class
