Imports System.Data.SqlClient
Imports System.data
Imports Management

Partial Class News_AddEditNews
    Inherits System.Web.UI.Page


    Protected Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.PreRender
        If Not Page.IsPostBack Then
            Me.Master.SetDisplayMessage("Please use the form below to edit the current news article.", Management.MessageType.GeneralMessage)
            Me.Master.PageTitle = "Edit News Information"
            Me.Master.PageHeader = "Edit News Information"
        End If

        If Request.QueryString("mode") = 1 Then
            Me.fvNews.DefaultMode = FormViewMode.Insert
        End If
    End Sub

    Protected Sub dsNews_Updated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsNews.Updated
        Response.Redirect("ManageNews.aspx")
    End Sub
    Protected Sub dsRecipe_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsNews.Inserted
        Response.Redirect("ManageNews.aspx")
    End Sub

    Protected Sub fvNews_ItemInserting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewInsertEventArgs) Handles fvNews.ItemInserting

        Try

            ' Get the FileUpload control
            Dim NewsDocument As FileUpload = CType(fvNews.FindControl("FileUpload"), FileUpload)

            If NewsDocument.PostedFile.FileName <> "" Then
                'Store the name of the new filename

                'Set the Save Path
                Dim savePath As String = BuildKnowledgeUploadPath(ConfigurationManager.AppSettings("CompanyID").ToString, ConfigurationManager.AppSettings("BrandID").ToString) + System.IO.Path.GetFileName(NewsDocument.PostedFile.FileName)
                e.Values("LocalSystemPath") = savePath

                'Upload File
                NewsDocument.SaveAs(savePath)
            Else
                'Store the name of the existing filename
                Dim strFileName As String = ""
                strFileName = System.IO.Path.GetFileName(CType(fvNews.FindControl("lblFile"), Label).Text)
                e.Values("LocalSystemPath") = strFileName

            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub fvNews_ItemUpdating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewUpdateEventArgs) Handles fvNews.ItemUpdating

        Try

            ' Get the FileUpload control
            Dim NewsDocument As FileUpload = CType(fvNews.FindControl("FileUpload"), FileUpload)

            If NewsDocument.PostedFile.FileName <> "" Then

                'Set the Save Path
                Dim savePath As String = BuildKnowledgeUploadPath(ConfigurationManager.AppSettings("CompanyID").ToString, ConfigurationManager.AppSettings("BrandID").ToString) + System.IO.Path.GetFileName(NewsDocument.PostedFile.FileName)

                'Store the name of the new filename
                e.NewValues("LocalSystemPath") = savePath

                'Upload File
                NewsDocument.SaveAs(savePath)
            Else
                'Store the name of the existing filename
                Dim strFileName As String = ""
                strFileName = CType(fvNews.FindControl("lblFile"), Label).Text
                'dsNews.UpdateParameters("LocalSystemPath").DefaultValue = strFileName
                e.NewValues("LocalSystemPath") = strFileName

            End If
        Catch ex As Exception

        End Try
    End Sub
    Public Function BuildKnowledgeUploadPath(ByVal Company As Integer, ByVal Brand As Integer) As String

        Dim strImageUploadPath As String = ""

        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

        Dim SQLCommand As New SqlCommand   'SQLCommand Object
        Dim SQLDataReader As SqlDataReader 'SQLDataReader
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Try 'Only one "Try" statement 

            'Set string to base image path (as a starting point)
            strImageUploadPath = ConfigurationManager.AppSettings("FileStorageLocation").ToString
            If Not System.IO.Directory.Exists(strImageUploadPath) Then
                System.IO.Directory.CreateDirectory(strImageUploadPath)
            End If
            SQLConn.Open() 'Open Database

            'Set the Basic Command Information 
            SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
            SQLCommand.Connection = SQLConn                          'Set the Connection

            SQLCommand.CommandText = "ZCMS..[GetResourceNamesForImageUploadPath]"

            'Set the ProdcutID Parameter
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Company))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Brand))

            'Set the DataReader
            SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)

            'Build path based on returned paramaters
            If SQLDataReader.Read Then
                strImageUploadPath = strImageUploadPath & SQLDataReader("Company") & "\" & SQLDataReader("Brand") & "\Knowledge\"
                If Not System.IO.Directory.Exists(strImageUploadPath) Then
                    System.IO.Directory.CreateDirectory(strImageUploadPath)
                End If
            End If

            'Cleanup
            SQLDataReader.Close()
            SQLConn.Close()

            BuildKnowledgeUploadPath = strImageUploadPath

        Catch SQLErr As SqlException
            BuildKnowledgeUploadPath = strImageUploadPath 'Returns just the base path in case of error
        Catch Err As Exception
            BuildKnowledgeUploadPath = strImageUploadPath 'Returns just the base path in case of error
        Finally
            'Confirm that The SQLDB Connection is closed
            If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
            SQLDataReader = Nothing
            SQLCommand = Nothing
            SQLConn = Nothing
        End Try

    End Function

End Class
