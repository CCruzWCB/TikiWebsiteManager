Imports Management


Public Class DataCube_EditDataCubePage
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

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'Put user code to initialize the page here
        If Not Page.IsPostBack Then
            'Capture referring Page
            Dim refPage As String = Request.ServerVariables("HTTP_REFERER")
            Response.Cookies("refPage").Value = refPage
            ViewState("DataCubeID") = Request.QueryString("DID").ToString
        Else
            If Not Len(Request.Cookies("refPage").Value.ToString) > 1 Then
                Response.Cookies("refPage").Value = "ViewDataCubeByPage.aspx"
            End If
        End If


        Dim oDataCube As New DataCube
        Dim sDataCube As sDateCubeElement
        Try

            If Len(ViewState("DataCubeID")) > 1 Then
                sDataCube = oDataCube.GetDataCubeElement(ViewState("DataCubeID"), False)
                If Not Len(sDataCube.DataCubeID) > 1 Then
                    Response.Redirect(Request.Cookies("refPage").Value)
                End If

                FormEditDataCube1.CurrentDataCubeType = sDataCube.TypeID
                FormEditDataCube1.CurrentDataCubeID = sDataCube.DataCubeID
                FormEditDataCube1.CurrentDataCube = sDataCube
            Else
                'Then Redirect back to select DataCube
                Response.Redirect(Request.Cookies("refPage").Value)
            End If
        Catch ex As Exception
            Response.Write("Error Loading DataCube")

        Finally
            oDataCube = Nothing

        End Try


    End Sub



End Class
