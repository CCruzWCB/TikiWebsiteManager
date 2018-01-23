Imports System.Data.SqlClient
'Imports BusinessLogic
Imports System.IO
'Imports TrickBag

Imports System.Data
Imports Management.TrickBag
Imports Management




Public Class WebControls_FormEditDataCube
    Inherits System.Web.UI.UserControl

#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        ' Dim configurationAppSettings As System.Configuration.AppSettingsReader = New System.Configuration.AppSettingsReader
        Me.sqlCnnFormEdit = New System.Data.SqlClient.SqlConnection
        '
        'sqlCnnFormEdit
        '
        Me.sqlCnnFormEdit.ConnectionString = ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ConnectionString
        'CType(configurationAppSettings.GetValue("BaseDBConnection", GetType(System.String)), String)

    End Sub
    'Form Elements and Controls
    'Protected WithEvents TableDataCube As System.Web.UI.WebControls.Table
    'Protected WithEvents btnCancel As System.Web.UI.WebControls.Button
    'Protected WithEvents btnSave As System.Web.UI.WebControls.Button
    'Protected WithEvents txtElement As System.Web.UI.WebControls.TextBox
    'Protected WithEvents lstType As System.Web.UI.WebControls.DropDownList
    'Protected WithEvents txtTargetUrl As System.Web.UI.WebControls.TextBox
    'Protected WithEvents lstTarget As System.Web.UI.WebControls.DropDownList
    'Protected WithEvents txtDescription As System.Web.UI.WebControls.TextBox
    'Protected WithEvents txtHeight As System.Web.UI.WebControls.TextBox
    'Protected WithEvents txtWidth As System.Web.UI.WebControls.TextBox
    'Protected WithEvents txtTargetPage As System.Web.UI.WebControls.TextBox
    'Protected WithEvents txtSourceCode As System.Web.UI.WebControls.TextBox
    'Protected WithEvents txtImagePath As System.Web.UI.WebControls.TextBox
    'Protected WithEvents ElementImage As System.Web.UI.WebControls.Image
    'Protected WithEvents lstEnable As System.Web.UI.WebControls.DropDownList
    'Protected WithEvents btnBrowseElements As System.Web.UI.WebControls.Button
    'Protected WithEvents txtOptionalText As System.Web.UI.WebControls.TextBox
    'Protected WithEvents imgMergeDescription As System.Web.UI.WebControls.Image
    'Protected WithEvents imageMergeOptionalText As System.Web.UI.WebControls.Image
    'Protected WithEvents hdnDataCubeID As System.Web.UI.HtmlControls.HtmlInputHidden
    'Protected WithEvents ImageFile As System.Web.UI.HtmlControls.HtmlInputFile
    'Protected WithEvents btnConfirm As System.Web.UI.WebControls.Button

    'Required Field Validators 
    'Protected WithEvents rfvType As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvElement As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvDescription As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvTargetUrl As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvTargetPage As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvTarget As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvOPtionalText As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvSourceCode As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvImageFile As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvImagePath As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvWidth As System.Web.UI.WebControls.RequiredFieldValidator
    'Protected WithEvents rfvHeight As System.Web.UI.WebControls.RequiredFieldValidator

    'Table Rows 
    'Protected WithEvents trHeader As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trElement As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trType As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trDescription As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trTargetUrl As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trTargetPage As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trTarget As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trOptionalText As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trSourceCode As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trDisplayImage As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trImage As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trImagePath As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trWidth As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trHeight As System.Web.UI.WebControls.TableRow
    'Protected WithEvents trEnable As System.Web.UI.WebControls.TableRow

    'Table Cells
    'Protected WithEvents tdTargetPageHeader As System.Web.UI.WebControls.TableCell

    'DataCubeControl 
    'Protected WithEvents DataCubePreview as  
    'Sql Connection 
    Protected WithEvents sqlCnnFormEdit As System.Data.SqlClient.SqlConnection

    'Panels
    'Protected WithEvents PanelEditMode As System.Web.UI.WebControls.Panel
    'Protected WithEvents PanelConfirm As System.Web.UI.WebControls.Panel
    'Protected WithEvents lblMessage As System.Web.UI.WebControls.Label



    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region

    Dim _iDataCubeType As Management.DataCubeType
    Dim _DataCubeID As String = String.Empty
    Dim _sDataCube As Management.sDateCubeElement
    Dim sDataSource As New Management.sDataSource
    Dim sDataCube As New Management.sDateCubeElement
    Dim _ErrorNumber As String = String.Empty
    Dim _ErrorDescription As String = String.Empty

    Public Const constFlashTemplate As String = "<object width={0} height={1} runat =""server"" id=""oFlash_{2}""" & _
                    "classid = ""clsid:d27cdb6e-ae6d-11cf-96b8-444553540000""" & _
                    "codebase = ""http://fpdownload.macromedia.com/pub/"" " & _
                    "shockwave/cabs/flash/swflash.cab#version=8,0,0,0"">" & _
                    "<param name=""movie"" value=""{3}"" />" & _
                    "<param name=""bgcolor"" value=""#ffffff"" />" & _
                    "<param name=""menu"" value=""false"" />" & _
                    "<param name=""quality"" value=""high"" />" & _
                    "<param name=""loop"" value=""false"" />" & _
                    "<embed src=""{3}"" width={0} height={1} " & _
                    "  bgcolor=""#ff0000"" menu=""false"" quality=""low"" " & _
                      "type=""application/x-shockwave-flash"" pluginspage=" & _
                      """http://www.macromedia.com/go/getflashplayer"" />" & _
                    "</object>"""

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here

        If Not Page.IsPostBack Then
            Me.sqlCnnFormEdit.ConnectionString = ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ConnectionString
            sDataCube = _sDataCube
            ViewState("DataCubeID") = sDataCube.DataCubeID
            InitPageProp(CurrentDataCubeType)
            SetFormData(CurrentDataCubeID)
            Me.PanelConfirm.Visible = False
            Me.PanelEditMode.Visible = True
            Me.btnSave.Visible = False
        Else
            sDataCube = _sDataCube
            If sDataCube.DataCubeID <> ViewState("DataCubeID") Then
                Response.Redirect(Request.Cookies("refPage").Value())
            End If
        End If
    End Sub

    Private Sub DisableAllFormControls()

        Me.rfvDescription.Enabled = False
        Me.rfvElement.Enabled = False
        Me.rfvHeight.Enabled = False
        Me.rfvHeight.Enabled = False
        Me.rfvImageFile.Enabled = False
        Me.rfvImagePath.Enabled = False
        Me.rfvOptionalText.Enabled = False
        Me.rfvSourceCode.Enabled = False ' Temporary Enabled
        Me.rfvTarget.Enabled = False
        Me.rfvTargetPage.Enabled = False
        Me.rfvTargetUrl.Enabled = False
        Me.rfvType.Enabled = False
        Me.rfvWidth.Enabled = False

        'newly added flash validators 
        Me.rfvFlashFile.Enabled = False

        'Enabled All Table Rows
        Me.trTargetPage.Visible = False
        Me.trTarget.Visible = False
        Me.trTargetUrl.Visible = False
        Me.trWidth.Visible = False
        Me.trHeader.Visible = False
        Me.trImage.Visible = False

        Me.trImagePath.Visible = False
        Me.trType.Visible = False
        Me.trOptionalText.Visible = False
        Me.trSourceCode.Visible = False
        Me.trDescription.Visible = False
        Me.trDisplayImage.Visible = False
        Me.trElement.Visible = False
        Me.trEnable.Visible = False
        Me.trHeight.Visible = False

        'newly added flash rows
        Me.trFlashFile.Visible = False
        Me.trDisplayFlash.Visible = False

        'Enabled Controls
        'Me.lstTarget.Visible = False
        'Me.lstType.Visible = False
        'Me.txtDescription.Visible = False
        'Me.txtImagePath.Visible = False
        'Me.trTargetPage.Visible = False
        'Me.txtTargetPage.Visible = False
        'Me.txtTargetUrl.Visible = False
        'Me.txtWidth.Visible = False
        'Me.txtHeight.Visible = False
        'Me.txtElement.Visible = False
        'Me.txtOptionalText.Visible = False
        'Me.txtSourceCode.Visible = False

    End Sub
    Private Sub InitPageProp(ByVal iDataCubeType As DataCubeType)
        'Disable all Elements and Controls
        DisableAllFormControls()

        Select Case iDataCubeType
            Case DataCubeType.Category

                'Enable validators
                Me.rfvElement.Enabled = True
                Me.rfvDescription.Enabled = True

                'Show table rows containing controls 
                Me.trHeader.Visible = True
                Me.trImage.Visible = True
                Me.trType.Visible = True
                Me.trDescription.Visible = True
                Me.trDisplayImage.Visible = True
                Me.trElement.Visible = True
                Me.trEnable.Visible = True

                'Show controls 
                Me.lstTarget.Visible = True
                Me.lstType.Visible = True
                Me.txtDescription.Visible = True
                Me.txtElement.Visible = True

                
                
            Case DataCubeType.Hyperlink

                'Set default values
                Me.tdTargetPageHeader.Text = "URL"


                'Show table rows containing controls 

                Me.trImage.Visible = True
                Me.trDescription.Visible = True
                Me.trTarget.Visible = True
                'Me.trTargetPage.Visible = True
                Me.trTargetUrl.Visible = True
                Me.trHeader.Visible = True
                Me.trType.Visible = True
                Me.trDisplayImage.Visible = True
                Me.trEnable.Visible = True


                'Show controls 
                Me.lstType.Visible = True
                Me.txtDescription.Visible = True
            Case DataCubeType.Image_Flash
                'Enable validators
                Me.rfvFlashFile.Enabled = True

                'Show table rows containing controls 
                Me.trHeader.Visible = True
                Me.trType.Visible = True
                Me.trFlashFile.Visible = True
                Me.trDisplayFlash.Visible = True
                Me.trEnable.Visible = True

                'Show controls 
                Me.lstType.Visible = True
                Me.divFlashContainer.Visible = True
                Me.FlashFile.Visible = True

            Case DataCubeType.Product
                'Enable validators
                Me.rfvElement.Enabled = True
                Me.rfvDescription.Enabled = True

                'Show table rows containing controls 
                Me.trHeader.Visible = True
                Me.trImage.Visible = True
                Me.trType.Visible = True
                Me.trDescription.Visible = True
                Me.trDisplayImage.Visible = True
                Me.trElement.Visible = True
                Me.trEnable.Visible = True

                'Show controls 
                Me.lstTarget.Visible = True
                Me.lstType.Visible = True
                Me.txtDescription.Visible = True
                Me.txtElement.Visible = True

            Case DataCubeType.ProductSeries
                'Enable validators
                Me.rfvElement.Enabled = True
                Me.rfvDescription.Enabled = True

                'Show table rows containing controls 
                Me.trHeader.Visible = True
                Me.trImage.Visible = True
                Me.trType.Visible = True
                Me.trDescription.Visible = True
                Me.trDisplayImage.Visible = True
                Me.trElement.Visible = True
                Me.trEnable.Visible = True

                'Show controls 
                Me.lstTarget.Visible = True
                Me.lstType.Visible = True
                Me.txtDescription.Visible = True
                Me.txtElement.Visible = True

            Case DataCubeType.Promotion
                'Enable validators
                Me.rfvElement.Enabled = True
                Me.rfvDescription.Enabled = True

                'Show table rows containing controls 
                Me.trHeader.Visible = True
                Me.trImage.Visible = True
                Me.trType.Visible = True
                Me.trDescription.Visible = True
                Me.trDisplayImage.Visible = True
                Me.trElement.Visible = True
                Me.trEnable.Visible = True

                'Show controls 
                Me.lstTarget.Visible = True
                Me.lstType.Visible = True
                Me.txtDescription.Visible = True
                Me.txtElement.Visible = True

            Case DataCubeType.Text_Image
                'Enable validators
                Me.rfvFlashFile.Enabled = True

                'Show table rows containing controls 
                Me.trHeader.Visible = True
                Me.trType.Visible = True
                Me.trImage.Visible = True
                Me.trDisplayImage.Visible = True
                Me.trEnable.Visible = True
                Me.trDescription.Visible = True

                'Show controls 
                Me.lstType.Visible = True
                Me.txtDescription.Visible = True


            Case DataCubeType.VirtualCategory
                'Enable validators
                Me.rfvElement.Enabled = True
                Me.rfvDescription.Enabled = True

                'Show table rows containing controls 
                Me.trHeader.Visible = True
                Me.trImage.Visible = True
                Me.trType.Visible = True
                Me.trDescription.Visible = True
                Me.trDisplayImage.Visible = True
                Me.trElement.Visible = True
                Me.trEnable.Visible = True

                'Show controls 
                Me.lstTarget.Visible = True
                Me.lstType.Visible = True
                Me.txtDescription.Visible = True
                Me.txtElement.Visible = True
        End Select

    End Sub

    Public Property CurrentDataCube() As Management.sDateCubeElement
        Get
            Return _sDataCube
        End Get
        Set(ByVal Value As sDateCubeElement)
            _sDataCube = Value
        End Set
    End Property

    Public Property CurrentDataCubeID() As String
        Get
            Return _DataCubeID
        End Get
        Set(ByVal Value As String)
            _DataCubeID = Value
        End Set
    End Property
    Public Property CurrentDataCubeType() As DataCubeType
        Get
            Return _iDataCubeType
        End Get
        Set(ByVal Value As DataCubeType)
            _iDataCubeType = Value
        End Set
    End Property

    Private Function GetElementOptions(ByVal sDataSource As sDataSource) As Data.DataSet
        'BuildSQL Statement
        Dim strTableName As String = ""
        Dim strWhereStatement As String = ""
        Dim strFieldList As String = ""
        Dim strSQL As String = ""

        'Build Select Statement
        If Not IsDBNull(sDataSource.TableDisplayField) Then
            If Not IsDBNull(sDataSource.TableValueField) Then
                strFieldList = "select " & sDataSource.TableDisplayField & " +'(' + " & sDataSource.TableValueField & " + ')' as [DisplayField] , " & sDataSource.TableValueField
            End If
        End If

        'Build From Statement 
        If Not IsDBNull(sDataSource.TableViewName) Then
            strTableName = " From " & sDataSource.TableViewName
        End If


        If Not (IsDBNull(sDataSource.TableKey1) Or IsNothing(sDataSource.TableKey1)) And Not (IsDBNull(sDataSource.SelectedElementID) Or IsNothing(sDataSource.SelectedElementID)) Then
            strWhereStatement = sDataSource.TableKey1 & " = '" & sDataSource.SelectedElementID & "'"
        End If
        If Len(strWhereStatement) > 1 Then
            strWhereStatement = " Where " & strWhereStatement
        End If

        strSQL = strFieldList & "" & strTableName & "" & strWhereStatement

        Dim oReader As SqlDataReader
        Dim oDataSet As New Data.DataSet

        Dim oCmd As New SqlCommand
        oCmd.CommandText = strSQL
        oCmd.Connection = Me.sqlCnnFormEdit
        If Me.sqlCnnFormEdit.State = Data.ConnectionState.Closed Then Me.sqlCnnFormEdit.Open()
        oReader = oCmd.ExecuteReader

        '  Dim oConvert As New Management.TrickBag.Convert
        'If oReader.Read Then
        oDataSet = Convert.DataReaderToDataSet(oReader)
        'End If

        Me.sqlCnnFormEdit.Close()

        Return oDataSet
    End Function

    Private Function GetMergeData(ByVal Element As String, ByVal FieldName As String, ByVal sDataSource As sDataSource) As String
        Dim isValid As Boolean = False
        Dim oDataCubeObject As New DataCube
        sDataSource = oDataCubeObject.GetDataSource(ViewState("DataCubeID"), Me.lstType.SelectedValue)
        ViewState("FieldValue") = sDataSource.TableValueField
        ViewState("TableDisplayField") = sDataSource.TableDisplayField
        ViewState("TableViewName") = sDataSource.TableViewName
        Dim strSQL As String
        Try
            If IsNumeric(Element) Then
                strSQL = "select Top 1 " & FieldName & " From " & sDataSource.TableViewName & " Where " & sDataSource.TableValueField & " = " & Element
            Else
                strSQL = "select Top 1 " & FieldName & " From " & sDataSource.TableViewName & " Where " & sDataSource.TableValueField & " = '" & Element & "'"
            End If

            Dim oCmd As New Data.SqlClient.SqlCommand
            oCmd.Connection = Me.sqlCnnFormEdit
            Me.sqlCnnFormEdit.Open()
            oCmd.CommandText = strSQL
            oCmd.CommandType = CommandType.Text

            GetMergeData = oCmd.ExecuteScalar

        Catch e As SqlException
            GetMergeData = FieldName
            lblMessage.Text = "Error Validating Element: " & e.Message.ToString
        Catch ex As Exception
            GetMergeData = FieldName
            lblMessage.Text = "Error Validating Element: " & ex.Message.ToString
        Finally
            If Me.sqlCnnFormEdit.State = ConnectionState.Open Then Me.sqlCnnFormEdit.Close()
        End Try

        Return GetMergeData

    End Function

    Private Function ValidateElement(ByVal Element As String) As Boolean
        Dim isValid As Boolean = False
        Dim oDataCubeObject As New Management.DataCube
        Dim sDataSource As Management.sDataSource
        sDataSource = oDataCubeObject.GetDataSource(ViewState("DataCubeID"), Me.lstType.SelectedValue)
        ViewState("FieldValue") = sDataSource.TableValueField
        ViewState("TableDisplayField") = sDataSource.TableDisplayField
        ViewState("TableViewName") = sDataSource.TableViewName
        Dim strSQL As String
        Try

            strSQL = "select count(*) From " & sDataSource.tableViewName & " Where " & sDataSource.tableValueField & " = '" & Element & "'"

            Dim oCmd As New SqlClient.SqlCommand
            oCmd.Connection = Me.sqlCnnFormEdit
            Me.sqlCnnFormEdit.Open()
            oCmd.CommandText = strSQL
            oCmd.CommandType = CommandType.Text

            Dim Exists As Integer = oCmd.ExecuteScalar
            If Exists > 0 Then isValid = True
        Catch e As SqlException
            isValid = False
            lblMessage.Text = "Error Validating Element: " & e.Message.ToString
        Catch ex As Exception
            isValid = False
            lblMessage.Text = "Error Validating Element: " & ex.Message.ToString
        Finally
            If Me.sqlCnnFormEdit.State = ConnectionState.Open Then Me.sqlCnnFormEdit.Close()
        End Try

        Return isValid

    End Function

    Private Function GetDataCubeTypes() As DataSet
        'BuildSQL Statement
        Dim strSQL As String

        'Build Select Statement

        strSQL = "Select DataItemType_Name , DataItemType_ID from DataItemType Where Active =1 Order By DataItemType_Name"

        Dim oReader As SqlDataReader
        Dim oDataSet As New DataSet

        Dim oCmd As New SqlCommand
        oCmd.CommandText = strSQL
        oCmd.Connection = Me.sqlCnnFormEdit
        If Me.sqlCnnFormEdit.State = ConnectionState.Closed Then Me.sqlCnnFormEdit.Open()
        oReader = oCmd.ExecuteReader

        '   Dim oConvert As New TrickBag.Convert
        'If oReader.Read Then
        oDataSet = Convert.DataReaderToDataSet(oReader)
        'End If

        Me.sqlCnnFormEdit.Close()

        Return oDataSet


    End Function
    Private Sub InitForm()
        Me.txtDescription.Text = ""
        Me.txtTargetUrl.Text = ""
        Me.txtTargetPage.Text = ""
        Me.lstTarget.SelectedValue = "_SELF"
        Me.txtElement.Text = ""
        Me.txtOptionalText.Text = ""
        Me.txtSourceCode.Text = ""
        Me.txtImagePath.Text = ""
        Me.txtWidth.Text = ""
        Me.txtHeight.Text = ""


    End Sub
    Private Sub SetFormData(ByVal DataCubeID As String)

        InitForm()

        'Store DataCube ID
        SetControlValue("hdnDataCubeID", DataCubeID)


        'Declare Objects
        Dim oDataSet As New DataSet
        Dim oDataCube As New Management.DataCube

        UpdateTableCell("tdHeader", True, DataCubeID)

        'Populate Type Dropdown
        sDataSource = oDataCube.GetDataSource(DataCubeID, 0)
        oDataSet = GetDataCubeTypes()
        Me.lstType.DataSource = oDataSet
        lstType.DataMember = oDataSet.Tables(0).TableName
        lstType.DataTextField = oDataSet.Tables(0).Columns(0).ColumnName ' "DisplayField"
        lstType.DataValueField = oDataSet.Tables(0).Columns(1).ColumnName 'item_id
        lstType.DataBind()

        Dim TempListItem3 = New ListItem
        With TempListItem3
            .Text = "Select A Type"
            .Value = 0
            .Selected = False
        End With

        lstType.Items.Insert(0, TempListItem3)

        If sDataCube.TypeID > 0 Then
            lstType.SelectedValue = sDataCube.TypeID
        End If

        If Len(sDataCube.ElementID) >= 1 Then
            Me.txtElement.Text = sDataCube.ElementID
        Else
            Me.txtElement.Text = "Enter or Select Element" 'Intial Value for Validator ---'"Enter or Select " & sDataCube.TableKey1
        End If

        If Len(sDataCube.TableKey1) > 1 Then
            UpdateTableCell("tdElementName", True, UCase(Replace(sDataCube.TableSelectionField, "_", " ")))
            ToggleTableRow("trElement", True)
            If CType(lstType.SelectedValue, DataItemType) = DataItemType.Product Then
                Me.btnBrowseElements.Attributes.Add("onclick", "ShowElements('" & DataCubeID & "'," & lstType.SelectedValue & ",'txtElement',400,350);")
                'Me.btnBrowseElements.Attributes.Add("onclick", "ShowItems('txtElement',400,350);")
            Else
                Me.btnBrowseElements.Attributes.Add("onclick", "ShowElements('" & DataCubeID & "'," & lstType.SelectedValue & ",'txtElement',400,350);")
            End If


            'Add Event for merge images
            imgMergeDescription.Attributes.Add("onclick", "MergeField('txtDescription', " & lstType.SelectedValue & ")")
            imageMergeOptionalText.Attributes.Add("onclick", "MergeField('txtOptionalText', " & lstType.SelectedValue & ")")
            imageMergeOptionalText.Visible = True
            imgMergeDescription.Visible = True
        Else
            ToggleTableRow("trElement", False)
            imageMergeOptionalText.Visible = False
            imgMergeDescription.Visible = False

        End If

        If sDataCube.AllowImageChange Then
            If Not IsDBNull(sDataCube.Image) Then
                If Len(sDataCube.Image) > 1 Then

                    If CType(lstType.SelectedValue, DataItemType) = DataItemType.Image_Flash Then
                        ToggleTableRow("trDisplayFlash", True) 'Show current Image 
                        divFlashContainer.Visible = True

                        'Set Flash properties 
                        Me.divFlashContainer.InnerHtml = String.Format(constFlashTemplate, sDataCube.Width, sDataCube.Height, DataCubeID, Utilities.GetImageURL(sDataCube.ImagePath & sDataCube.Image))

                        
                    ElseIf CType(lstType.SelectedValue, DataItemType) <> DataItemType.Image_Flash Then
                        If Not sDataCube.Image.Contains(".swf") Then
                            'Not a flash file
                            Me.ElementImage.ImageUrl = Utilities.GetImageURL(sDataCube.ImagePath & sDataCube.Image)
                            Me.ElementImage.Width = Unit.Pixel(sDataCube.Width)
                            Me.ElementImage.Height = Unit.Pixel(sDataCube.Height)
                            ToggleTableRow("trDisplayImage", True) 'Show current Image 
                        Else
                            ToggleTableRow("trDisplayImage", False) 'Show current Image 
                        End If
                    Else

                    End If

                Else
                    Me.ElementImage.ImageUrl = ""
                    ToggleTableRow("trDisplayImage", False) 'Show current Image 
                    ToggleTableRow("trDisplayFlash", False) 'Show current Image 
                End If
            End If

            'ToggleTableRow("trDisplayImage", True) 'Show current Image 
            'ToggleTableRow("trImage", True) 'Show Image Browse Button 
            'ToggleTableRow("trTarget", True)
        Else
            'ToggleTableRow("trDisplayImage", False) 'HideImage Browse Button 
            If CType(lstType.SelectedValue, DataItemType) = DataItemType.Image_Flash Or CType(lstType.SelectedValue, DataItemType) = DataItemType.Text_Image Or CType(lstType.SelectedValue, DataItemType) = DataItemType.Text Then
                ToggleTableRow("trImage", False)  'Hide Image Browse Button 
                ToggleTableRow("trFlashFile", False)  'Hide Image Browse Button 
                ToggleTableRow("trDisplayImage", False) 'hidecurrent Image 
            End If

            'ToggleTableRow("trHeight", False)
            'ToggleTableRow("trWidth", False)

        End If

        Try


            Me.txtDescription.Text = sDataCube.Description
            Me.txtTargetUrl.Text = sDataCube.ScriptPage_URL
            Me.txtTargetPage.Text = sDataCube.ScriptPage_Name
            Me.lstTarget.SelectedValue = UCase(sDataCube.Target)

            Me.txtOptionalText.Text = sDataCube.Optional_Text
            Me.txtSourceCode.Text = sDataCube.SourceCode
            Me.txtImagePath.Text = sDataCube.ImagePath
            Me.txtWidth.Text = sDataCube.Width
            Me.txtHeight.Text = sDataCube.Height

            Me.lstEnable.SelectedValue = Math.Abs(CType(sDataCube.Enabled, Integer))

        Catch ex As Exception
            Me.lblMessage.Text = "Error Loading values: " & ex.Message.ToString
        Finally
            oDataCube = Nothing
        End Try






    End Sub

    Private Sub SetControlValue(ByVal ControlName As String, ByVal strValue As String) ', ByVal Type As ControlType)

        Dim tobject As Object = Me.FindControl(ControlName)
        If tobject.GetType().Name = "HtmlInputHidden" Then

        End If
        Try
            Select Case tobject.GetType().Name
                Case "HtmlInputText"
                    tobject.value = strValue
                Case "HtmlInputHidden"
                    tobject.value = strValue
                Case "Label"
                    tobject.text = strValue
                Case "DropDownList"
                    tobject.selectedValue = strValue
                Case "TextBox"
                    tobject.text = strValue
            End Select

        Catch ex As Exception

        End Try
        tobject = Nothing

    End Sub
    Private Sub UpdateTableCell(ByVal CellID As String, ByVal blnShowRow As Boolean, Optional ByVal NewTextValue As String = "")
        Dim td As TableCell
        td = Me.TableDataCube.FindControl(CellID)
        Try
            If NewTextValue <> "" Then
                td.Text = NewTextValue
            End If
            td.Visible = blnShowRow
        Catch ex As Exception
        Finally
            td = Nothing
        End Try
    End Sub
    Private Sub ToggleTableRow(ByVal RowID As String, ByVal blnShowRow As Boolean)
        Dim tr As TableRow
        tr = Me.TableDataCube.FindControl(RowID)
        Try
            tr.Visible = blnShowRow
        Catch ex As Exception
        Finally
            tr = Nothing
        End Try
    End Sub
    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Response.Redirect(Request.Cookies("refPage").Value)
    End Sub
    Private Function GetImageName(ByVal FullImageName As String) As String
        Dim tempImageName As String = ""
        Dim pos As Integer = 0
        Dim ilen As Integer = 0


        pos = InStrRev(FullImageName, "/")
        If pos = 0 Then
            pos = InStrRev(FullImageName, "\")
        End If

        ilen = Len(FullImageName) - pos
        Return Right(FullImageName, ilen)

    End Function
    Private Sub ConfirmUpdate()

        Me.PanelConfirm.Visible = True
        Me.PanelEditMode.Visible = False
        Me.btnConfirm.Visible = False
        Me.btnSave.Visible = True
        Dim sPreviewDataCube As Management.sDateCubeElement = Nothing

        Dim bFileUpload As Boolean = False
        DataCubePreview.Enabled = True
        sPreviewDataCube.Enabled = True

        DataCubePreview.DataCubeType = CType(lstType.SelectedValue, DataCubeType)
        sPreviewDataCube.TypeID = CType(lstType.SelectedValue, DataCubeType)
        DataCubePreview.DataCubeID = ViewState("DataCubeID")
        sPreviewDataCube.DataCubeID = ViewState("DataCubeID")
        DataCubePreview.ElementID = Trim(Me.txtElement.Text)
        sPreviewDataCube.ElementID = Trim(Me.txtElement.Text)
        DataCubePreview.TextBody = Me.txtDescription.Text
        sPreviewDataCube.Description = Me.txtDescription.Text
        DataCubePreview.TargetUrl = Me.txtTargetUrl.Text
        sPreviewDataCube.ScriptPage_URL = Me.txtTargetUrl.Text

        DataCubePreview.TargetPage = Me.txtTargetPage.Text
        sPreviewDataCube.ScriptPage_Name = Me.txtTargetPage.Text

        DataCubePreview.Target = Me.lstTarget.SelectedValue
        sPreviewDataCube.Target = Me.lstTarget.SelectedValue

        DataCubePreview.QueryString = Me.txtTargetUrl.Text & Me.txtTargetPage.Text
        DataCubePreview.SourceCode = Me.txtSourceCode.Text

        Select Case CType(lstType.SelectedValue, DataCubeType)
            Case DataCubeType.Image_Flash 'Flash
                If Not Me.FlashFile.PostedFile Is Nothing AndAlso Me.FlashFile.PostedFile.ContentLength > 0 Then
                    Dim FlashFile As String = GetImageName(Me.FlashFile.PostedFile.FileName) 'ToDo function should return "" if empty
                    If Len(FlashFile) > 1 Then
                        DataCubePreview.ImageName = FlashFile
                        DataCubePreview.ImagePath = ConfigurationManager.AppSettings("ImageStagingDirectory").ToString
                        'Get Upload Path 
                        If Not Directory.Exists(DataCubePreview.ImagePath) Then
                            Try
                                Directory.CreateDirectory(DataCubePreview.ImagePath)
                            Catch ex As Exception

                            End Try

                        End If

                        Dim UploadPath As String = ConfigurationManager.AppSettings("ImageStagingDirectory").ToString 'GetMappedPath(ConfigurationSettings.AppSettings("ImageStagingDirectory").ToString)
                        Me.FlashFile.PostedFile.SaveAs(UploadPath & FlashFile)    'Try to save path.
                        bFileUpload = True
                        'Store previous ImageName and Upload Image 
                        DataCubePreview.ImageName = FlashFile
                    End If
                Else

                    DataCubePreview.ImageName = GetImageName(Me.ElementImage.ImageUrl)
                    sPreviewDataCube.Image = GetImageName(Me.ElementImage.ImageUrl)
                End If


                ViewState("ImageName") = DataCubePreview.ImageName

                If Len(DataCubePreview.ImageName) > 1 Then
                    If bFileUpload Then
                        DataCubePreview.ImageUrl = ConfigurationManager.AppSettings("ImageStagingDirectory").ToString() & DataCubePreview.ImageName
                    Else
                        DataCubePreview.ImageUrl = Me.txtImagePath.Text & DataCubePreview.ImageName
                        sPreviewDataCube.Image = Me.txtImagePath.Text & DataCubePreview.ImageName
                    End If

                End If
            Case DataCubeType.Text_Image, DataCubeType.Hyperlink, DataCubeType.Product, DataCubeType.ProductSeries, DataCubeType.Promotion, DataCubeType.Category



                If Not Me.ImageFile.PostedFile Is Nothing AndAlso Me.ImageFile.PostedFile.ContentLength > 0 Then
                    Dim ImageName As String = GetImageName(Me.ImageFile.PostedFile.FileName) 'ToDo function should return "" if empty
                    If Len(ImageName) > 1 Then
                        DataCubePreview.ImageName = ImageName
                        DataCubePreview.ImagePath = ConfigurationManager.AppSettings("ImageStagingDirectory").ToString
                        'Get Upload Path 
                        If Not Directory.Exists(DataCubePreview.ImagePath) Then
                            Directory.CreateDirectory(DataCubePreview.ImagePath)
                        End If

                        Dim UploadPath As String = ConfigurationManager.AppSettings("ImageStagingDirectory").ToString 'GetMappedPath(ConfigurationSettings.AppSettings("ImageStagingDirectory").ToString)
                        Me.ImageFile.PostedFile.SaveAs(UploadPath & ImageName)    'Try to save path.
                        bFileUpload = True
                        'Store previous ImageName and Upload Image 
                        DataCubePreview.ImageName = ImageName
                    End If
                Else

                    DataCubePreview.ImageName = GetImageName(Me.ElementImage.ImageUrl)
                    sPreviewDataCube.Image = GetImageName(Me.ElementImage.ImageUrl)
                End If


                ViewState("ImageName") = DataCubePreview.ImageName

                If Len(DataCubePreview.ImageName) > 1 Then
                    If bFileUpload Then
                        DataCubePreview.ImageUrl = ConfigurationManager.AppSettings("ImageStagingDirectory").ToString() & DataCubePreview.ImageName
                    Else
                        DataCubePreview.ImageUrl = Me.txtImagePath.Text & DataCubePreview.ImageName
                        sPreviewDataCube.Image = Me.txtImagePath.Text & DataCubePreview.ImageName
                    End If

                End If
            Case Else

        End Select





        DataCubePreview.Height = Me.txtHeight.Text
        DataCubePreview.Width = Me.txtWidth.Text

        DataCubePreview.PreviewDataCube(ViewState("DataCubeID"), sPreviewDataCube)
        DataCubePreview.QueryString = sPreviewDataCube.ScriptPage_URL & sPreviewDataCube.ScriptPage_Name
        DataCubePreview.TextBody = sPreviewDataCube.Description

        DataCubePreview.BuildDataCube()


    End Sub
    Private Sub btnSave_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave.Click

        Dim TypeID As String = lstType.SelectedValue
        sDataCube.TypeID = TypeID

        Dim ElementID As String = txtElement.Text
        sDataCube.ElementID = ElementID
        Dim Description As String = txtDescription.Text
        sDataCube.Description = Description

        Dim TargetUrl As String = txtTargetUrl.Text
        sDataCube.ScriptPage_URL = TargetUrl
        Dim TargetPage As String = txtTargetPage.Text
        sDataCube.ScriptPage_Name = TargetPage
        Dim Target As String = lstTarget.SelectedValue
        sDataCube.Target = Target
        Dim OptionalText As String = txtOptionalText.Text
        sDataCube.Optional_Text = OptionalText
        Dim SourceCode As String = txtSourceCode.Text
        sDataCube.SourceCode = SourceCode
        Dim ImagePath As String = txtImagePath.Text
        If Right(ImagePath, 1) <> "/" Then
            'Append ending slash /
            ImagePath += "/"
        End If
        sDataCube.ImagePath = ImagePath
        Dim Element_Image As String = GetImageName(Me.ElementImage.ImageUrl)
        sDataCube.Image = Element_Image

        Dim DataCubeID As String = hdnDataCubeID.Value
        sDataCube.DataCubeID = DataCubeID


        'Copy File from Temp Location to Final Location
        Dim StagingPath As String = ConfigurationManager.AppSettings("ImageStagingDirectory").ToString
        If Len(ViewState("ImageName")) > 1 Then
            Try


                If IO.File.Exists(StagingPath & ViewState("ImageName")) Then
                    Dim UploadPath As String = ConfigurationManager.AppSettings("DataCubeImageDirectory")
                    If Not Directory.Exists(UploadPath) Then
                        Directory.CreateDirectory(UploadPath)
                    End If
                    IO.File.Copy(StagingPath & ViewState("ImageName"), UploadPath & ViewState("ImageName"), True)
                    IO.File.Delete(StagingPath & ViewState("ImageName"))
                    sDataCube.Image = Utilities.GetImageURL(ViewState("ImageName"))
                    sDataCube.ImagePath = UploadPath

                End If
            Catch ex As Exception
                UpdateTableCell("tdHeader", True, ex.Message.ToString)
                Exit Sub
            End Try

        End If
        Dim Width As String = txtWidth.Text
        sDataCube.Width = Width
        Dim Height As String = txtHeight.Text
        sDataCube.Height = Height
        Dim Enable As Integer = lstEnable.SelectedValue
        sDataCube.Enabled = CType(Enable, Boolean)

        Dim oDataCube As New Management.DataCube
        oDataCube.Update(sDataCube)
        sDataCube = oDataCube.GetDataCubeElement(sDataCube.DataCubeID, False)
        SetFormData(sDataCube.DataCubeID)
        oDataCube = Nothing

        Response.Redirect(Request.Cookies("refPage").Value)
    End Sub

    Private Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Unload
        Me.sqlCnnFormEdit = Nothing
        GC.Collect()
    End Sub

    Private Sub lstType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstType.SelectedIndexChanged

        Dim oDataCube As New Management.DataCube
        Dim TypeID As String = Me.lstType.SelectedValue
        If TypeID > 0 Then
            sDataCube = oDataCube.GetDataCubeDefaultByType(ViewState("DataCubeID"), TypeID, False)
        End If
        InitPageProp(CType(TypeID, DataCubeType))
        SetFormData(sDataCube.DataCubeID)
        oDataCube = Nothing
    End Sub

    Private Function GetMappedPath(ByVal VirtualDir As String) As String
        Dim strpos As Integer = 0
        Dim endpos As Integer = 0
        Dim length As Integer = Len(VirtualDir)
        Dim temp As String = ""

        If InStr(VirtualDir, "//") > 0 Then 'Convert to virtual dir path
            strpos = InStr(VirtualDir, "//")
            VirtualDir = Right(VirtualDir, length - (strpos + 2))
            endpos = InStr(VirtualDir, "/")
            length = Len(VirtualDir)
            temp = Mid(VirtualDir, endpos)
            temp = Server.MapPath(temp)
        End If

        Return temp

    End Function

    Private Sub btnConfirm_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnConfirm.Click
        'Validate Select Element
        If Me.txtElement.Visible = True Then
            If ValidateElement(Me.txtElement.Text) Then
                ConfirmUpdate()
            Else
                lblMessage.Text = "Element Select is Invalid"
            End If
        Else
            ConfirmUpdate()
        End If


    End Sub
End Class
