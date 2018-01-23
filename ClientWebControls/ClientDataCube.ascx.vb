Imports System.Web
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports Management
Imports System.Data


Namespace ClientDataCubeControl

    Partial Class ClientWebControls_ClientDataCube
        Inherits System.Web.UI.UserControl

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
                        "</object>"

        Dim _Name As String = String.Empty
        Dim _Index As Integer
        Dim _DataCubeID As String = String.Empty
        Dim _ClassName As String = String.Empty
        Dim _TextColor As Drawing.Color = Drawing.Color.Black
        Dim _BorderColor As Drawing.Color = Drawing.Color.Black
        Dim _BorderSize As Integer = 1
        Dim _EnableBorder As Boolean = False
        Dim _Enabled As Boolean = False
        Dim _SourcePage As String = String.Empty
        Dim _ElementID As String = String.Empty
        Dim _Height As Integer = 0
        Dim _Width As Integer = 0
        Dim _ImageName As String = String.Empty
        Dim _ImagePath As String = String.Empty
        Dim _ImageUrl As String = String.Empty
        Dim _ActiveDate As String = ""
        Dim _InActiveDate As String = ""
        Dim _TargetPage As String = String.Empty
        Dim _TargetUrl As String = String.Empty
        Dim _Target As String = ""
        Dim _CategoryID As Integer = 0
        Dim _PromotionID As Integer = 0
        Dim _ProductSeriesID As Integer = 0
        Dim _TextBody As String = String.Empty
        Dim _SourceCode As String = String.Empty
        Dim _OptionalText As String = String.Empty
        Dim _EditMode As Boolean = False
        Dim _ErrorNumber As String = String.Empty
        Dim _ErrorDescription As String = String.Empty
        Dim _QueryString As String = String.Empty
        Dim _DataCubeType As BPS_BL.BPS.DataCubeType



        Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
            'Put user code to initialize the page here

            If Not Page.IsPostBack Then
                Init_DataCube()
            End If

        End Sub

        'Properties of DataCube Control 

#Region " Properties"
        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal Value As String)
                _Name = Value
            End Set
        End Property
        Public Property DataCubeID() As String
            Get
                Return _DataCubeID
            End Get
            Set(ByVal Value As String)
                _DataCubeID = Value
            End Set
        End Property
        Public Property ClassName() As String
            Get
                Return _ClassName
            End Get
            Set(ByVal Value As String)
                _ClassName = Value
            End Set
        End Property
        Public Property TextColor() As Drawing.Color
            Get
                Return _TextColor
            End Get
            Set(ByVal Value As Drawing.Color)
                _TextColor = Value

            End Set
        End Property
        Public Property BorderColor() As Drawing.Color
            Get
                Return _BorderColor
            End Get
            Set(ByVal Value As Drawing.Color)
                _BorderColor = Value
            End Set
        End Property
        Public Property BorderSize() As Integer
            Get
                Return _BorderSize
            End Get
            Set(ByVal Value As Integer)
                _BorderSize = Value
            End Set
        End Property
        Public Property EnableBorder() As Boolean
            Get
                Return _EnableBorder
            End Get
            Set(ByVal Value As Boolean)
                _EnableBorder = Value
            End Set
        End Property
        Public Property Enabled() As Boolean
            Get
                Return _Enabled
            End Get
            Set(ByVal Value As Boolean)
                _Enabled = Value
            End Set
        End Property

        Public Property EditMode() As Boolean
            Get
                Return _EditMode
            End Get
            Set(ByVal Value As Boolean)
                _EditMode = Value
            End Set
        End Property
        Public Property SourcePage() As String
            Get
                Return _SourcePage
            End Get
            Set(ByVal Value As String)
                _SourcePage = Value
            End Set
        End Property
        Public Property ElementID() As String
            Get
                Return _ElementID
            End Get
            Set(ByVal Value As String)
                _ElementID = Value
            End Set
        End Property
        Public Property Height() As Integer
            Get
                Return _Height
            End Get
            Set(ByVal Value As Integer)
                _Height = Value
            End Set
        End Property
        Public Property Width() As Integer
            Get
                Return _Width
            End Get
            Set(ByVal Value As Integer)
                _Width = Value
            End Set
        End Property
        Public Property ImageName() As String
            Get
                Return _ImageName
            End Get
            Set(ByVal Value As String)
                _ImageName = Value
            End Set
        End Property
        Public Property ImagePath() As String
            Get
                Return _ImagePath
            End Get
            Set(ByVal Value As String)
                _ImagePath = Value
            End Set
        End Property
        Public Property ImageUrl() As String
            Get
                Return _ImageUrl
            End Get

            Set(ByVal Value As String)
                _ImageUrl = Value

            End Set
        End Property
        Public Property ActiveDate() As String
            Get
                Return _ActiveDate
            End Get
            Set(ByVal Value As String)
                _ActiveDate = Value
            End Set
        End Property
        Public Property InActiveDate() As String
            Get
                Return _InActiveDate
            End Get
            Set(ByVal Value As String)
                _InActiveDate = Value
            End Set
        End Property
        Public Property TargetPage() As String
            Get
                Return _TargetPage
            End Get
            Set(ByVal Value As String)
                _TargetPage = Value
            End Set
        End Property
        Public Property TargetUrl() As String
            Get
                Return _TargetUrl
            End Get
            Set(ByVal Value As String)
                _TargetUrl = Value

            End Set
        End Property
        Public Property Target() As String
            Get
                Return _Target
            End Get
            Set(ByVal Value As String)
                _Target = Value
            End Set
        End Property
        Public Property TextBody() As String
            Get
                Return _TextBody
            End Get
            Set(ByVal Value As String)
                _TextBody = Value

            End Set
        End Property
        Public Property SourceCode() As String
            Get
                Return _SourceCode
            End Get
            Set(ByVal Value As String)
                _SourceCode = Value
            End Set
        End Property
        Public Property OptionalText() As String
            Get
                Return _OptionalText
            End Get
            Set(ByVal Value As String)
                _OptionalText = Value
            End Set
        End Property
        Public Property QueryString() As String
            Get
                Return _QueryString
            End Get
            Set(ByVal Value As String)
                _QueryString = Value
            End Set
        End Property

        Public Property DataCubeType() As Management.DataCubeType
            Get
                Return _DataCubeType
            End Get
            Set(ByVal Value As Management.DataCubeType)
                _DataCubeType = CType(Value, Management.DataCubeType)
            End Set
        End Property

        Public Property Index() As Integer
            Get
                Return _Index
            End Get
            Set(ByVal value As Integer)
                _Index = value
            End Set
        End Property


        Public Property CategoryID() As Integer
            Get
                Return _CategoryID
            End Get
            Set(ByVal value As Integer)
                _CategoryID = value
                Init_DataCube()
            End Set
        End Property

        Public Property PromotionID() As Integer
            Get
                Return _PromotionID
            End Get
            Set(ByVal value As Integer)
                _PromotionID = value
                Init_DataCube()
            End Set
        End Property
        Public Property ProductSeriesID() As Integer
            Get
                Return _ProductSeriesID
            End Get
            Set(ByVal value As Integer)
                _ProductSeriesID = value
                Init_DataCube()
            End Set
        End Property

        Public ReadOnly Property ErrorNumber()
            Get
                Return _ErrorNumber
            End Get

        End Property
        Public ReadOnly Property ErrorDescription()
            Get
                Return _ErrorDescription
            End Get

        End Property
#End Region

        'Methods of DataCube Control 

        Public Sub LoadDataCube(ByVal DataCubeID As String)
            Me._DataCubeID = DataCubeID
            Init_DataCube()
        End Sub

        Public Sub LoadDataCube(ByVal UniqueID As Integer, ByVal Index As Integer, ByVal DataCubeType As DataCubeType)
            If DataCubeType = Management.DataCubeType.Category Then
                Me._CategoryID = UniqueID
                Me._Index = Index
            ElseIf DataCubeType = Management.DataCubeType.Promotion Then
                Me._PromotionID = UniqueID
                Me._Index = Index
            ElseIf DataCubeType = Management.DataCubeType.ProductSeries Then
                Me._ProductSeriesID = UniqueID
                Me._Index = Index
            End If
            Init_DataCube()
        End Sub
       
        Private Sub Init_DataCube()
            'Initialize Control with value from DataCubeObject
            Dim oDataCube As New Management.DataCube
            Dim sDataItemElement As Management.sDateCubeElement = Nothing
            Dim DataCubeCategoryID As String
            Dim DataCubePromotionID As String
            Dim DataCubeSeriesID As String
            Dim bDataCubeFound As Boolean = False
            Dim bDataCubeCreationError As Boolean = False


            If _CategoryID > 0 And _Index > 0 Then ' Get DataCube for specified Category


                'Use Category Prefix to identify the name of datacube to be 
                'shown based on category id
                DataCubeCategoryID = ConfigurationManager.AppSettings("DataCubeCategoryNamePrefix") & _CategoryID & "_" & _Index
                sDataItemElement = oDataCube.GetDataCubeElement(DataCubeCategoryID)
            ElseIf _PromotionID > 0 And _Index > 0 Then ' Get DataCube for specified Promotion
                'Use Promotion Prefix to identify the name of datacube to be 
                'shown based on Promotion id
                DataCubePromotionID = ConfigurationManager.AppSettings("DataCubePromotionNamePrefix") & _PromotionID & "_" & _Index
                sDataItemElement = oDataCube.GetDataCubeElement(DataCubePromotionID)
            ElseIf _ProductSeriesID > 0 And _Index > 0 Then ' Get DataCube for specified Promotion
                'Use Promotion Prefix to identify the name of datacube to be 
                'shown based on Promotion id
                DataCubeSeriesID = ConfigurationManager.AppSettings("DataCubeProductSeriesNamePrefix") & _PromotionID & "_" & _Index
                sDataItemElement = oDataCube.GetDataCubeElement(DataCubeSeriesID)
            Else
                If _DataCubeID <> "" Then
                    sDataItemElement = oDataCube.GetDataCubeElement(_DataCubeID)
                    If IsNothing(sDataItemElement.DataCubeID) Then
                        If _EditMode Then
                            'This must be a new category ID, create new datacube
                            If Not oDataCube.CreateDefaultDataCube(_DataCubeID, DataCubeType.Category) Then
                                bDataCubeCreationError = True

                            End If
                        End If
                    End If
                End If
            End If

            If sDataItemElement.Enabled Or _EditMode Then
                  _DataCubeType = sDataItemElement.TypeID
                _Name = sDataItemElement.Name
                _SourcePage = sDataItemElement.SourcePage
                _ElementID = sDataItemElement.ElementID
                _ImageName = sDataItemElement.Image
                _ImagePath = sDataItemElement.ImagePath
                _ActiveDate = sDataItemElement.ActiveDate
                _InActiveDate = sDataItemElement.InActiveDate
                _TargetPage = sDataItemElement.ScriptPage_Name
                _TargetUrl = sDataItemElement.ScriptPage_URL
                _Enabled = sDataItemElement.Enabled

                _OptionalText = sDataItemElement.Optional_Text

                '************************************
                ''Check for Overrideable Properties
                '************************************
                '   _ClassName = Set by Developer
                '   _TextColor = Set by Developer
                '   _BorderColor = Set by Developer
                '   _BorderSize = Set by Developer
                '   _EnabledBorder = Set by Developer
                '   _QueryString = Set by Developer

                If Not Len(_DataCubeID) > 0 Then _DataCubeID = sDataItemElement.DataCubeID
                If Not Len(_QueryString) > 0 Then _QueryString = _TargetUrl & _TargetPage 'Todo Check for empty string and /
                If Not Len(_SourceCode) > 0 Then _SourceCode = sDataItemElement.SourceCode
                If Not IsNothing(_ImagePath) And Not IsNothing(_ImageName) > 0 Then
                    _ImageUrl = _ImagePath & _ImageName 'Todo Check for empty string and /
                Else
                    _ImageUrl = Utilities.GetImageURL(ConfigurationManager.AppSettings("NoDataCubeImageURL"))
                End If

                If Not Len(_TextBody) > 0 Then _TextBody = sDataItemElement.Description
                If Not Len(_Target) > 0 Then _Target = sDataItemElement.Target
                If Not _Height > 0 Then _Height = sDataItemElement.Height
                If Not _Width > 0 Then _Width = sDataItemElement.Width

                If _EditMode Then 'Build QueryString 
                    If Not bDataCubeCreationError Then
                        If _DataCubeType = BPS_BL.BPS.DataCubeType.Text Then
                            'don't lose the text value
                            _QueryString = "../DataCube/EditDataCubePage.aspx?DID=" & DataCubeID
                            '_TextBody = "Click to Edit DataCube " & DataCubeID  'functions as alt text when used with image. 
                            _OptionalText = "Click to Edit DataCube " & DataCubeID  'functions as alt text when used with image. 
                        Else
                            _QueryString = "../DataCube/EditDataCubePage.aspx?DID=" & DataCubeID
                            _TextBody = "Click to Edit DataCube " & DataCubeID  'functions as alt text when used with image. 
                            _OptionalText = ""
                        End If
                        
                    Else

                        _TextBody = "There was a problem generating DataCube " & DataCubeID & ". Please contact your web site hosting provider" 'functions as alt text when used with image. 
                        _OptionalText = ""
                    End If

                End If

                sDataItemElement = Nothing
                oDataCube = Nothing

                BuildDataCube(_DataCubeID)
            Else
                Me.DataLink.Visible = False
                'Me.DataImage.Visible = False
            End If

        End Sub


        Public Sub BuildDataCube(Optional ByVal DataCubeID As String = "")
            Try

                '*********************************
                'Set Default Image 
                '*********************************
                Dim bIsFlashFile As Boolean = False

                '*********************************
                'Set Default Image 
                '*********************************

                'Check to see if file ext is a flash file 
                If ImageUrl.Contains(".swf") Then
                    bIsFlashFile = True
                End If

                If bIsFlashFile Then

                    '*********************************
                    'Set Flash File Properties 
                    '*********************************
                    divFlashContainer.Visible = True

                    Me.divFlashContainer.InnerHtml = String.Format(constFlashTemplate, Width, Height, DataCubeID, Replace(Utilities.GetImageURL(ImageUrl), ConfigurationManager.AppSettings("ApplicationHTTPRoot"), ""))

                    If _EditMode = True Then


                        '*********************************
                        'Set Default QueryString
                        '*********************************
                        If Not IsDBNull(QueryString) Then
                            '**************************
                            'Append SourceCode to QueryString
                            '**************************
                            If Len(QueryString) > 1 Then
                                If Len(SourceCode) > 1 Then
                                    QueryString = Trim(QueryString)
                                    If InStr(QueryString, "?") > 0 Then
                                        QueryString += "&sc=" & SourceCode  'Todo Set variable Name for source querystring param
                                    Else
                                        QueryString += "?sc=" & SourceCode  'Todo Set variable Name for source querystring param
                                    End If
                                End If
                            End If

                            If Len(QueryString) > 1 Then
                                Me.DataLink.Text = "Click here to Edit Flash"
                                Me.DataLink.NavigateUrl = QueryString
                            End If
                        End If
                    End If
                Else
                    '*********************************
                    'Must be a standard image or other element type 
                    '*********************************
                    If _DataCubeType <> BPS_BL.BPS.DataCubeType.Text Then


                        Me.divFlashContainer.InnerHtml = ""
                        divFlashContainer.Visible = False

                        If Not IsDBNull(ImageUrl) Then



                            Me.DataLink.ImageUrl = Utilities.GetImageURL(ImageUrl)


                            'Me.DataImage.Attributes.Add("usemap", "#" & Left(ImageName, Len(ImageName) - 4) & "_Map")
                            'Me.DataImage.ImageUrl = ImageUrl
                            'Me.DataImage.Attributes.Add("hspace", "0")
                            'Me.DataImage.Attributes.Add("vspace", "0")
                            ' Me.DataLink.Controls(0).Attributes.Add("width", Width)
                            Me.DataLink.Attributes.Add("height", Height)

                            If EnableBorder Then
                                '   Me.DataImage.BorderStyle = BorderStyle.Solid
                                '   Me.DataImage.BorderColor = BorderColor
                                '   Me.DataImage.BorderWidth = Unit.Pixel(BorderSize)
                            End If

                            'If Not IsDBNull(Width) Then Me.DataImage.Width = Unit.Pixel(Width)
                            'If Not IsDBNull(Height) Then Me.DataImage.Height = Unit.Pixel(Height)

                            DataLink.Width = Unit.Pixel(Width)
                            DataLink.Height = Unit.Pixel(Height)
                        End If
                    Else
                        Me.DataLink.ImageUrl = ""
                    End If


                    '*********************************
                    'Set Default Text Properties (Used as HyperLink Text or Tooltip
                    '*********************************
                    If Len(_OptionalText) > 1 Then
                        Me.DataLink.Text = OptionalText 'functions as alt text when used with image. 
                        Me.DataLink.ToolTip = OptionalText
                    Else

                        Me.DataLink.Text = TextBody 'functions as alt text when used with image. 
                        Me.DataLink.ToolTip = TextBody
                    End If

                    '*********************************
                    'Set Default QueryString
                    '*********************************
                    If Not IsDBNull(QueryString) Then
                        '**************************
                        'Append SourceCode to QueryString
                        '**************************
                        If Len(QueryString) > 1 Then
                            If Len(SourceCode) > 1 Then
                                QueryString = Trim(QueryString)
                                If InStr(QueryString, "?") > 0 Then
                                    QueryString += "&sc=" & SourceCode  'Todo Set variable Name for source querystring param
                                Else
                                    QueryString += "?sc=" & SourceCode  'Todo Set variable Name for source querystring param
                                End If
                            End If
                        End If

                        If Len(QueryString) > 1 Then
                            Me.DataLink.NavigateUrl = QueryString
                        End If
                    End If

                    '*********************************
                    'Set Page Target Attribute
                    '*********************************
                    If UCase(Target) = "_BLANK" Then
                        Me.DataLink.Target = Target
                    Else
                        'Default is _SELF  DO NOT SET 
                    End If

                End If



                '*********************************
                'If Edit Mode - OverRide text properties 
                '*********************************
                If _EditMode = True Then

                    Me.DataLink.BorderStyle = BorderStyle.Solid
                    Me.DataLink.BorderWidth = Unit.Pixel(5)

                    Me.DataLink.ForeColor = Drawing.Color.Black

                    If Not Me._Enabled Then
                        Me.DataLink.BorderColor = Drawing.Color.Red
                    Else
                        Me.DataLink.BorderColor = Drawing.Color.Yellow
                    End If
                    '   Me.DataImage.BorderStyle = BorderStyle.Solid
                    '   Me.DataImage.BorderColor = BorderColor
                    '   Me.DataImage.BorderWidth = Unit.Pixel(BorderSize)
                End If

                If _Enabled = True Or _EditMode = True Then

                    Me.DataLink.Visible = True
                    'Me.DataImage.Visible = False
                    If EnableBorder Then
                        If Len(ImageUrl) < 1 And EditMode = False Then 'Border should be on Image only
                            'Apply border to <a tag if not image set
                            Me.DataLink.BorderStyle = BorderStyle.Solid
                            Me.DataLink.BorderWidth = Unit.Pixel(BorderSize)
                            Me.DataLink.BorderColor = BorderColor
                        End If
                    End If

                    Me.DataLink.ForeColor = TextColor

                    If Len(ClassName) > 1 Then
                        Me.DataLink.CssClass = ClassName
                    Else
                        'Me.DataLink.CssClass = "classes_small"
                        'Me.DataLink.Style.Add("text-align", "center") 'Need vertical alignment 
                        Me.DataLink.Style.Add("text-decoration", "none")
                        Me.DataLink.Style.Add("text-align", "center")
                    End If

                    If _DataCubeType = BPS_BL.BPS.DataCubeType.Text Then
                        Me.DataText.Visible = True
                        Me.DataText.InnerHtml = TextBody
                        Me.DataText.Style.Add("width", Width.ToString)
                        Me.DataText.Style.Add("height", Height.ToString)

                        'Me.DataLink.Visible = False
                    End If

                Else
                    Me.DataLink.Visible = False
                    'Me.DataImage.Visible = False
                End If

            Catch exhp As HttpParseException
                _ErrorNumber = exhp.GetHttpCode
                _ErrorDescription = exhp.Message.ToString
                Me.DataLink.Visible = False
                'Me.DataImage.Visible = False

            Catch exh As HttpException
                _ErrorNumber = exh.ErrorCode
                _ErrorDescription = exh.Message.ToString
                Me.DataLink.Visible = False
                'Me.DataImage.Visible = False

            Catch ex As Exception
                _ErrorNumber = -1
                _ErrorDescription = ex.Message.ToString
                Me.DataLink.Visible = False
                'Me.DataImage.Visible = False

            End Try

        End Sub

        Public Function PreviewDataCube(ByVal DataCubeID As String, ByVal sDataCubePreviewElement As Management.sDateCubeElement) As Management.sDateCubeElement
            Dim oCnn As New SqlConnection(ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ConnectionString)
            Dim oCmd As New SqlCommand
            Dim oReader As SqlDataReader
            Dim oDataReader As SqlDataReader
            Dim sDataItemElement As New Management.sDateCubeElement
            Dim strQueryStringParams As String = ""
            Dim strTempQueryStringParam As String = ""
            Dim bHasMergeField As Boolean = False
            Dim strFieldList As String = ""

            Try
                oCnn.Open()

                oCmd.CommandText = "GetDataCubeTypeParam"
                oCmd.CommandType = System.Data.CommandType.StoredProcedure
                oCmd.Connection = oCnn

                oCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DataCubeID", System.Data.SqlDbType.VarChar, 50, Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", Data.DataRowVersion.Current, DataCubeID))

                oReader = oCmd.ExecuteReader()

                Dim RowCount As Integer = 0

                While oReader.Read()
                    With sDataItemElement


                        'Get List of all table Seletion Params
                        If Not (Microsoft.VisualBasic.IsDBNull(oReader("table_View_Name"))) Then sDataItemElement.tableViewName = oReader("table_View_Name")
                        If Not (Microsoft.VisualBasic.IsDBNull(oReader("table_SelectionField"))) Then sDataItemElement.tableSelectionField = oReader("table_SelectionField")

                        'Get List of all parameters. 
                        If Not (Microsoft.VisualBasic.IsDBNull(oReader("ParamName")) AndAlso Microsoft.VisualBasic.IsDBNull(oReader("ParamValue"))) Then
                            strTempQueryStringParam += oReader("ParamName") & "=" & oReader("ParamValue") & "&"
                            'Check For Mergable Data Fields in querystring parameters
                            If InStr(oReader("ParamValue"), "<@") AndAlso InStr(oReader("ParamValue"), "@>") Then
                                bHasMergeField = True
                                If Not InStr(strFieldList, GetFieldName(oReader("ParamValue"))) > 0 Then
                                    strFieldList += GetFieldName(oReader("ParamValue")) & ","
                                End If

                            End If
                            strQueryStringParams = strQueryStringParams & "" & strTempQueryStringParam
                            strTempQueryStringParam = ""
                        End If

                        RowCount += 1
                    End With
                End While

                If Len(strQueryStringParams) > 1 Then
                    'Trim Last Character from QueryString
                    strQueryStringParams = Left(strQueryStringParams, Len(strQueryStringParams) - 1)
                End If


                oCmd = Nothing
                oReader.Close()


                'Check for an data that should be populated from database table...The data is contain within specified delimiters. <@item_id@>

                'For the Description Field Check for multible mergeable fields
                Dim StrArray As Array = Split(sDataCubePreviewElement.Description, " ")
                Dim z As Integer = 0
                For z = 0 To StrArray.Length - 1
                    If InStr(StrArray(z), "<@") AndAlso InStr(StrArray(z), "@>") Then
                        If Not InStr(strFieldList, GetFieldName(StrArray(z))) > 0 Then
                            strFieldList += GetFieldName(StrArray(z)) & ","
                            bHasMergeField = True
                        End If
                    End If
                Next
                StrArray = Nothing
                'If InStr(sDataItemElement.Description, "<@") AndAlso InStr(sDataItemElement.Description, "@>") Then
                '    bHasMergeField = True
                '    If Not InStr(strFieldList, GetFieldName(sDataItemElement.Description)) > 0 Then
                '        strFieldList += GetFieldName(sDataItemElement.Description) & ","
                '    End If
                'End If

                If InStr(sDataCubePreviewElement.Image, "<@") AndAlso InStr(sDataCubePreviewElement.Image, "@>") Then
                    bHasMergeField = True
                    If Not InStr(strFieldList, GetFieldName(sDataCubePreviewElement.Image)) > 0 Then
                        strFieldList += GetFieldName(sDataCubePreviewElement.Image) & ","
                    End If
                End If

                If InStr(sDataCubePreviewElement.ImagePath, "<@") AndAlso InStr(sDataCubePreviewElement.ImagePath, "@>") Then
                    bHasMergeField = True
                    If Not InStr(strFieldList, GetFieldName(sDataCubePreviewElement.ImagePath)) > 0 Then
                        strFieldList += GetFieldName(sDataCubePreviewElement.ImagePath) & ","
                    End If
                End If

                If InStr(sDataCubePreviewElement.ScriptPage_Name, "<@") AndAlso InStr(sDataCubePreviewElement.ScriptPage_Name, "@>") Then
                    bHasMergeField = True
                    If Not InStr(strFieldList, GetFieldName(sDataCubePreviewElement.ScriptPage_Name)) > 0 Then
                        strFieldList += GetFieldName(sDataCubePreviewElement.ScriptPage_Name) & ","
                    End If
                End If

                If InStr(sDataCubePreviewElement.Optional_Text, "<@") AndAlso InStr(sDataCubePreviewElement.Optional_Text, "@>") Then
                    bHasMergeField = True
                    If Not InStr(strFieldList, GetFieldName(sDataCubePreviewElement.Optional_Text)) > 0 Then
                        strFieldList += GetFieldName(sDataCubePreviewElement.Optional_Text) & ","
                    End If

                End If

                If InStr(sDataCubePreviewElement.SourceCode, "<@") AndAlso InStr(sDataCubePreviewElement.SourceCode, "@>") Then
                    bHasMergeField = True
                    If Not InStr(strFieldList, GetFieldName(sDataCubePreviewElement.SourceCode)) > 0 Then
                        strFieldList += GetFieldName(sDataCubePreviewElement.SourceCode) & ","
                    End If

                End If

                If InStr(sDataCubePreviewElement.ScriptPage_URL, "<@") AndAlso InStr(sDataCubePreviewElement.ScriptPage_URL, "@>") Then
                    bHasMergeField = True
                    If Not InStr(strFieldList, GetFieldName(sDataCubePreviewElement.ScriptPage_URL)) > 0 Then
                        strFieldList += GetFieldName(sDataCubePreviewElement.ScriptPage_URL) & " ,"
                    End If

                End If





                'Make sure sql string is valid 
                If bHasMergeField Then
                    'Prepare to Merge Data Fields.
                    strFieldList = Left(strFieldList, Len(strFieldList) - 1)
                    Dim strtableName As String = sDataItemElement.tableViewName
                    Dim strWhereClause As String = ""
                    Dim ParamName As String = "@" & sDataItemElement.tableSelectionField
                    strWhereClause = sDataItemElement.tableSelectionField & " = " & ParamName

                    Try
                        If oCnn.State = Data.ConnectionState.Closed Then oCnn.Open()
                        oCmd = New SqlCommand
                        oCmd.CommandText = String.Format("Select {0} from {1} Where {2} = {3}", strFieldList, strtableName, sDataItemElement.tableSelectionField, ParamName)
                        oCmd.CommandType = System.Data.CommandType.Text
                        oCmd.Connection = oCnn
                        oCmd.Parameters.Add(New System.Data.SqlClient.SqlParameter(ParamName, System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, RTrim(sDataCubePreviewElement.ElementID)))
                        oDataReader = oCmd.ExecuteReader()

                        If oDataReader.Read Then
                            'Replace merge field names with data
                            Dim i As Integer = 0
                            Do Until i > oDataReader.FieldCount() - 1

                                Dim strValue As String = ""
                                If IsDBNull(oDataReader(i)) Then
                                    strValue = ""
                                Else
                                    strValue = CType(oDataReader(i), String)
                                End If
                                sDataItemElement.Description = Replace(sDataCubePreviewElement.Description, "<@" & oDataReader.GetName(i) & "@>", strValue)
                                sDataItemElement.Image = Replace(sDataCubePreviewElement.Image, "<@" & oDataReader.GetName(i) & "@>", strValue)
                                sDataItemElement.ImagePath = Replace(sDataCubePreviewElement.ImagePath, "<@" & oDataReader.GetName(i) & "@>", strValue)
                                sDataItemElement.Optional_Text = Replace(sDataCubePreviewElement.Optional_Text, "<@" & oDataReader.GetName(i) & "@>", strValue)
                                sDataItemElement.SourceCode = Replace(sDataCubePreviewElement.SourceCode, "<@" & oDataReader.GetName(i) & "@>", strValue)
                                sDataItemElement.ScriptPage_Name = Replace(sDataCubePreviewElement.ScriptPage_Name, "<@" & oDataReader.GetName(i) & "@>", strValue)
                                sDataItemElement.ScriptPage_URL = Replace(sDataCubePreviewElement.ScriptPage_URL, "<@" & oDataReader.GetName(i) & "@>", strValue)
                                strQueryStringParams = Replace(strQueryStringParams, "<@" & oDataReader.GetName(i) & "@>", strValue)
                                i += 1
                            Loop
                            'Todo If merge fails remove all field names <@fieldname@>
                        End If


                        sDataItemElement.ScriptPage_Name += "?" & strQueryStringParams

                        'sDataItemElement.ImagePath = Utilities.GetImageURL(sDataItemElement.ImagePath)




                    Catch ex As SqlException
                        _ErrorDescription = ex.Message.ToString
                    Catch e As Exception
                        _ErrorDescription = e.Message.ToString
                    End Try

                End If

            Catch ex As SqlException
                _ErrorDescription = ex.Message.ToString

            Catch e As Exception
                _ErrorDescription = e.Message.ToString
            Finally
                If oCnn.State = Data.ConnectionState.Open Then oCnn.Close()
                oCnn = Nothing
            End Try


            Return sDataItemElement

            sDataItemElement = Nothing


        End Function

        Private Function GetFieldName(ByVal strFieldName As String) As String
            'Check for an data that should be populated from database table...The data is contain within specified delimiters. <@item_id@>
            'Strip on delimiters and return field name
            Dim strPos As Integer
            Dim endPos As Integer
            Dim tempFieldName As String
            Dim length As Integer

            strPos = InStr(strFieldName, "<@")
            endPos = InStr(strFieldName, "@>")

            'Dim oReg As New Regex("<@

            length = endPos - strPos
            length = length - 2

            tempFieldName = Mid(strFieldName, strPos + 2, length)



            Return tempFieldName




        End Function


    End Class
End Namespace