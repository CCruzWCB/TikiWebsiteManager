Imports System.Web
Imports System.Data
Imports management


Public Class DataCube_mSelectMergeFields
    Inherits System.Web.UI.Page

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Dim sDataSource As New sDataSource
    Dim oDataCubeObject As New DataCube
    Dim controlName As String


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        'Put user code to initialize the page here


        If Not Page.IsPostBack Then
            Dim strDataCubeID As String = Request.QueryString("did")
            Dim strTypeID As Integer = CType(Request.QueryString("tid"), Integer)
            controlName = Request.QueryString("ctr")
            ViewState("controlName") = controlName
            sDataSource = oDataCubeObject.GetDataSource("test", strTypeID)
            PopulateFieldName(sDataSource.TypeID)
        Else

        End If



    End Sub

    Private Sub PopulateFieldName(ByVal DataCubeType As DataCubeType)
        Dim ConnectionString As String = ""
        Dim CommandText As String = ""


        'Check to see which database should be checked. 
        Select Case DataCubeType
            'Check BPS Database for these items 
            Case Management.DataCubeType.Category, Management.DataCubeType.Product, Management.DataCubeType.ProductSeries, Management.DataCubeType.VirtualCategory
                ConnectionString = ConfigurationManager.ConnectionStrings("BaseDBConnection").ConnectionString
                'Check CB web site Database for these items 
                CommandText = "[GetColumnNames]"
            Case Management.DataCubeType.Promotion
                ConnectionString = ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ConnectionString
                CommandText = "DataCube.[GetColumnNames]"
            Case Else
                ConnectionString = ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ConnectionString
                CommandText = "DataCube.[GetColumnNames]"
        End Select


        Dim oCnn As New Data.SqlClient.SqlConnection(ConnectionString)

        Dim oCmd As New SqlClient.SqlCommand
        Dim oDataSet As New DataSet

        oCmd.CommandText = CommandText
        oCmd.CommandType = CommandType.StoredProcedure
        oCmd.Connection = oCnn

        oCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TableViewName", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sDataSource.tableViewName))

        Try
            oCnn.Open()
            Dim oReader As SqlClient.SqlDataReader
            oReader = oCmd.ExecuteReader
            oDataSet = TrickBag.Convert.DataReaderToDataSet(oReader)

            Me.dgFieldName.DataSource = oDataSet
            Me.dgFieldName.DataMember = oDataSet.Tables(0).TableName
            'Me.dgFieldName.DataKeyField = oDataSet.Tables(0).Columns(0).ColumnName ' "DisplayField"
            Me.dgFieldName.DataBind()

        Catch ex As Exception
            Response.Write("error " & ex.Message.ToString)
        Finally
            If oCnn.State = ConnectionState.Open Then oCnn.Close()
        End Try



    End Sub

    Private Sub dgFieldName_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles dgFieldName.ItemCommand
        If e.CommandName = "add" Then
            Dim test As String = e.Item.Cells(1).Text()

            Dim hdnControl As UI.HtmlControls.HtmlInputHidden
            hdnControl = Me.FindControl("hdnColumn")
            hdnControl.Value = "<@" & test & "@>"

            Dim hdnControlName As UI.HtmlControls.HtmlInputHidden
            hdnControlName = Me.FindControl("hdnControlName")
            hdnControlName.Value = ViewState("controlName")


        End If


    End Sub
End Class
