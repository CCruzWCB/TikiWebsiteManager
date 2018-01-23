Imports System.Web
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports management.DataCube
Imports Management
Imports System.Drawing


Partial Class HTMLControls_DataCubeControl
    Inherits System.Web.UI.UserControl

    Dim _Name As String = String.Empty
    Dim _DataCubeID As String = String.Empty
    Dim _ClassName As String = String.Empty
    Dim _TextColor As Color = Color.Black
    Dim _BorderColor As Color = Color.Black
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
    Dim _TextBody As String = String.Empty
    Dim _SourceCode As String = String.Empty
    Dim _OptionalText As String = String.Empty
    Dim _EditMode As Boolean = False
    Dim _ErrorNumber As String = String.Empty
    Dim _ErrorDescription As String = String.Empty
    Dim _QueryString As String = String.Empty
    Dim _DataCubeType As Management.DataCubeType


    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here

        Init_DataCube()
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
    Public Property TextColor() As Color
        Get
            Return _TextColor
        End Get
        Set(ByVal Value As Color)
            _TextColor = Value

        End Set
    End Property
    Public Property BorderColor() As Color
        Get
            Return _BorderColor
        End Get
        Set(ByVal Value As Color)
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

    Public Property DataCubeType() As DataCubeType
        Get
            Return _DataCubeType
        End Get
        Set(ByVal Value As DataCubeType)
            _DataCubeType = CType(Value, DataCubeType)
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
    Private Sub Init_DataCube()
        'Initialize Control with value from DataCubeObject
        Dim oDataCube As New Management.DataCube
        Dim sDataItemElement As New Management.sDateCubeElement
        '            sDataItemElement = oDataCube.GetDataCubeDefaultByType(DataCubeID, DataCubeType, True)
        sDataItemElement = oDataCube.GetDataCubeElement(DataCubeID)
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

        If Not Len(_DataCubeID) > 0 Then _DataCubeID = sDataItemElement.DataCubeID
        If Not Len(_QueryString) > 0 Then _QueryString = _TargetUrl & _TargetPage 'Todo Check for empty string and /
        If Not Len(_SourceCode) > 0 Then _SourceCode = sDataItemElement.SourceCode
        If Not Len(_ImageUrl) > 0 Then
            _ImageUrl = _ImagePath & _ImageName 'Todo Check for empty string and /
        Else
            '  _ImageUrl = _ImagePath & _ImageName 'Todo Check for empty string and /
        End If

        If Not Len(_TextBody) > 0 Then _TextBody = sDataItemElement.Description
        If Not Len(_Target) > 0 Then _Target = sDataItemElement.Target
        If Not _Height > 0 Then _Height = sDataItemElement.Height
        If Not _Width > 0 Then _Width = sDataItemElement.Width




        sDataItemElement = Nothing
        oDataCube = Nothing

        If Enabled Or _EditMode Then
            BuildDataCube()
        Else
            Me.DataLink.Visible = False
            Me.DataImage.Visible = False
        End If

    End Sub
    Public Sub BuildDataCube()
        Try

            '*********************************
            'Set Default Image 
            '*********************************

            If Not IsDBNull(ImageUrl) Then
                If Len(ImageUrl) > 0 Then

                    If Len(ImageName) > 1 Then
                        Me.DataImage.Attributes.Add("usemap", "#" & Left(ImageName, Len(ImageName) - 4) & "_Map")
                        Me.DataImage.ImageUrl = ImageUrl
                        Me.DataImage.Attributes.Add("hspace", "0")
                        Me.DataImage.Attributes.Add("vspace", "0")
                        'Me.DataImage.Attributes.Add("width", Width)
                        'Me.DataImage.Attributes.Add("height", Height)
                        If EnableBorder Then
                            Me.DataImage.BorderStyle = BorderStyle.Solid
                            Me.DataImage.BorderColor = BorderColor
                            Me.DataImage.BorderWidth = Unit.Pixel(BorderSize)
                        End If
                        If Not IsDBNull(Width) Then Me.DataImage.Width = Unit.Pixel(Width)
                        If Not IsDBNull(Height) Then Me.DataImage.Height = Unit.Pixel(Height)
                    Else
                        ImageUrl = ""
                    End If

                End If
            End If


            '*********************************
            'Set Default Text Properties (Used as HyperLink Text or Tooltip
            '*********************************
            If Len(ImageUrl) > 1 Then
                Me.DataImage.AlternateText = TextBody 'functions as alt text when used with image. 
            Else
                If Len(OptionalText) > 1 Then
                    Me.DataLink.Text = OptionalText 'functions as alt text when used with image. 
                    Me.DataLink.ToolTip = OptionalText
                Else
                    Me.DataLink.Text = TextBody 'functions as alt text when used with image. 
                    Me.DataLink.ToolTip = TextBody
                End If
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


            If Enabled = True Or EditMode = True Then
                Me.DataLink.Visible = True
                Me.DataImage.Visible = True
                If EnableBorder Then
                    If Len(ImageUrl) < 1 And EditMode = False Then 'Border should be on Image only
                        'Apply border to <a tag if not image set
                        Me.DataLink.BorderStyle = BorderStyle.Solid
                        Me.DataLink.BorderWidth = Unit.Pixel(BorderSize)
                        Me.DataLink.BorderColor = BorderColor
                    ElseIf EditMode = True And Len(ImageUrl) < 1 Then
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
            Else
                Me.DataLink.Visible = False
                Me.DataImage.Visible = False
            End If

        Catch exh As HttpException
            _ErrorNumber = exh.ErrorCode
            _ErrorDescription = exh.Message.ToString
            Me.DataLink.Visible = False
            Me.DataImage.Visible = False

        Catch ex As Exception
            _ErrorNumber = -1
            _ErrorDescription = ex.Message.ToString
            Me.DataLink.Visible = False
            Me.DataImage.Visible = False

        End Try

    End Sub

End Class

