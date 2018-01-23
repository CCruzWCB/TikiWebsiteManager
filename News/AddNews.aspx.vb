Imports System.Data.SqlClient
Imports System.Data

Partial Class Media_AddNews
    Inherits System.Web.UI.Page


    Protected Sub rdoContentType_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoContentType.SelectedIndexChanged

        TogglePanels(Me.rdoContentType.SelectedValue)

    End Sub

    Private Sub TogglePanels(ByVal ContentTypeID As Integer)

        'SEt all to false 
        Me.pHtmlTExt.Visible = False
        Me.pURL.Visible = False
        Me.pUpload.Visible = False

        'Activate selected panel
        Select Case rdoContentType.SelectedValue
            Case "1" 'HTML /TEXT 
                Me.pHtmlTExt.Visible = True
            Case "2" 'URL 
                Me.pURL.Visible = True
            Case "3" 'FILE 
                Me.pUpload.Visible = True

        End Select
    End Sub
    Protected Sub btnAddNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddNew.Click
        Dim bsuccess As Boolean = False

        Select Case Me.rdoContentType.SelectedValue
            Case "1" 'HTML/TExt 
                bsuccess = AddNews(1, "", "")
            Case "2" 'URL 
                bsuccess = AddNews(2, Me.txtURL.Text, "")
            Case "3" 'File 
                'Upload and process the new file. 
                bsuccess = UploadPDFFile(Me.fileUpload)
        End Select

        If bsuccess Then
            Response.Redirect("Managenews.aspx?Message=News Item was Added sucessfully!&MessageType=" & Management.MessageType.ConfirmationMessage)
        End If




    End Sub

    Public Function AddNews(ByVal ContentTypeID As Integer, Optional ByVal URL As String = "", Optional ByVal LocalSystemPath As String = "") As Boolean

        Dim bSuccess As Boolean 'Return variable 

        'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ToString)

        Dim SQLCommand As New SqlCommand    'SQLCommand Object
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '*********************************************************************************************
        ' ****************** To Add Knowledge One Step is Required: ********************
        '   - 1. Add the Knowledge table entry
        '   
        '*********************************************************************************************

        Try    'Only one "Try" statement 

            SQLConn.Open()    'Open Database


            'Set the Basic Command Information 
            SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
            SQLCommand.Connection = SQLConn       'Set the Connection


            'Set the Specific Command Information 
            SQLCommand.CommandText = "LLFBPS.dbo.[ManagerAddWebsiteNews]"     'Stored Procedure Name

            'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Title", System.Data.SqlDbType.VarChar, 250)).Value = Me.txtTitle.Text
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeCategoryID", System.Data.SqlDbType.VarChar, 250)).Value = Me.ddCategory.SelectedValue
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.Text)).Value = Me.txtArticle.Text
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ContentTypeID", System.Data.SqlDbType.Int, 4)).Value = Me.rdoContentType.SelectedValue
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CanPublishInternal", System.Data.SqlDbType.Bit, 1))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1)).Value = Me.chkActive.Checked
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 150))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@LocalSystemPath", System.Data.SqlDbType.VarChar, 500))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


            If IsNumeric(Page.User.Identity.Name) Then
                SQLCommand.Parameters("@CreatedBy").Value = Page.User.Identity.Name
            Else
                SQLCommand.Parameters("@CreatedBy").Value = 0
            End If

            SQLCommand.Parameters("@CreatedName").Value = "Website Administrator"

            'Check to see if we should store Uploaded Path or HTTP Path
            Select Case ContentTypeID
                Case 1 'html 
                    SQLCommand.Parameters("@LocalSystemPath").Value = ""
                    SQLCommand.Parameters("@CanPublishInternal").Value = 1

                Case 2 'url 
                    SQLCommand.Parameters("@LocalSystemPath").Value = URL
                    SQLCommand.Parameters("@CanPublishInternal").Value = CBool(Me.rdoContentType.SelectedValue)
                Case 3 'file 
                    SQLCommand.Parameters("@LocalSystemPath").Value = LocalSystemPath
                    SQLCommand.Parameters("@CanPublishInternal").Value = 0 ' force all files to appear in new window.

            End Select



            SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

            If Not SQLCommand.Parameters("@KnowledgeID").Value > 0 Then
                bSuccess = False
            Else
                bSuccess = True
                ' Me.Master.SetDisplayMessage("News Item was added successfully", Management.MessageType.ConfirmationMessage)
                ResetForm(1) 'Reset back to HTML/text 
            End If

            'Close Connection
            SQLConn.Close()

            'SQL ERROR CATCH
        Catch SQLErr As SqlException
            bSuccess = False


            'MISC ERROR CATCH
        Catch Err As Exception
            bSuccess = False

        Finally
            'Confirm that The SQLDB Connection is closed
            If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
            SQLConn = Nothing
            SQLCommand = Nothing

        End Try

        'Set returning Boolean
        Return bSuccess



    End Function

    Public Function BuildKnowledgeUploadPath(ByVal Company As Integer, ByVal Brand As Integer) As String

        Dim strImageUploadPath As String = ""

        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("LLFPublicWebsiteConnectionString").ToString) 'SQLConnection Object

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

            SQLCommand.CommandText = "LLFBPS.dbo.[GetResourceNamesForImageUploadPath]"

            'Set the ProdcutID Parameter
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Company))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Brand))

            'Set the DataReader
            SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)

            'Build path based on returned paramaters
            If SQLDataReader.Read Then
                strImageUploadPath = strImageUploadPath & SQLDataReader("Company") & "\" & SQLDataReader("Brand") & "\Knowledge\NewsContent\"
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

    Private Function UploadPDFFile(ByRef oFileUPload As FileUpload, Optional ByRef UploadedFilePath As String = "") As Boolean
        Dim blnHasAttachment As Boolean
        Dim strSaveLocation As String
        Dim strFilename As String
        Dim blnFileExist As Boolean = True

        Dim bSuccess As Boolean

        Try

            'First Upload the Attachment (If Necessary)
            If Not oFileUPload.PostedFile Is Nothing And oFileUPload.PostedFile.ContentLength > 0 Then
                blnHasAttachment = True
                'strSaveLocation = ConfigurationSettings.AppSettings("FileStorageLocation")  'Used to Save a uploaded File
                strSaveLocation = BuildKnowledgeUploadPath(ConfigurationManager.AppSettings("CompanyID").ToString, ConfigurationManager.AppSettings("BrandID").ToString)

                strFilename = System.IO.Path.GetFileName(oFileUPload.PostedFile.FileName)   'Used when you have to rename a file 

                If LCase(strFilename) Like "*pdf*" Then
                    'Check to see if file with Samename already Exist
                    Do Until blnFileExist = False
                        If System.IO.File.Exists(strSaveLocation & strFilename) Then
                            blnFileExist = True
                            strFilename = String.Format("{0}_{1}_{2}", System.IO.Path.GetFileNameWithoutExtension(strFilename), Left(System.Guid.NewGuid.ToString, 10), System.IO.Path.GetExtension(strFilename))
                        Else
                            blnFileExist = False
                        End If
                    Loop

                    'Upload & Save the File
                    oFileUPload.PostedFile.SaveAs(strSaveLocation & strFilename)

                    'set uploaded file apth 
                    UploadedFilePath = (strSaveLocation & strFilename)

                    'Add the File and save to database
                    If AddNews(3, "", strSaveLocation & strFilename) Then
                        bSuccess = True
                    Else
                        System.IO.File.Delete(strSaveLocation & strFilename)
                        Me.Master.SetDisplayMessage("Error uploading file", Management.MessageType.ErrorMessage)

                    End If

                Else
                    bSuccess = False

                    Me.Master.SetDisplayMessage("Incorrect file format.  The selected file must be a PDF file.  Please check your file and try again.", Management.MessageType.ErrorMessage)
                End If
            Else
                bSuccess = False

                Me.Master.SetDisplayMessage("No File Selected.  Please select a file and try again.", Management.MessageType.SyntaxMessage)
            End If

        Catch ex As Exception
            Me.Master.SetDisplayMessage(ex.Message, Management.MessageType.ErrorMessage)
        Finally

        End Try

        Return bSuccess

    End Function

    Private Sub ResetForm(ByVal ContentTypeID As Integer)
        'Clear form and prepare for next ad
        Me.txtTitle.Text = ""
        Me.txtArticle.Text = ""
        Me.rdoContentType.SelectedValue = ContentTypeID
        Me.chkActive.Checked = False
        TogglePanels(ContentTypeID) ' Reset back to HTML/Text options. 
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            Me.Master.PageHeader = "Add News Item"
            Me.Master.PageTitle = "Add News Items"
            Me.Master.SetDisplayMessage("Complete the form below to Add a new News Item.", Management.MessageType.GeneralMessage)


        End If
    End Sub
End Class
