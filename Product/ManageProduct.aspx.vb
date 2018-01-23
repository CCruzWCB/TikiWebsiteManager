Imports System.Windows.Forms.Application
Imports BPS_BL.BPS
Imports Management
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Data.SqlClient
Imports System.Data

Partial Class Manager_ManageProduct
    Inherits System.Web.UI.Page
    Dim imageHeight As Integer = 0
    Dim imageWidth As Integer = 0
    Dim ImagePath As String = ""
    Dim TempUploadPath As String = ""
    Dim NewFileName As String = ""
    Dim FullImagePath As String = ""
    Dim CompanyID As Integer = 0
    Dim BrandID As Integer = 0
    Dim ResourceTypeID As Integer = 0

    Protected Sub Page_Disposed(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Disposed
        'Clean up any base image uploaded to the temp directory 
        CleanUpOldImages()
    End Sub
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
     

        Me.Master.PageHeader = "Manage Product"
        Me.Master.PageTitle = "Manage Product"

        If Not Page.IsPostBack Then
            Try

                Me.Master.SetDisplayMessage("Use the form below to manage product.", Management.MessageType.GeneralMessage)


                smMain.RegisterAsyncPostBackControl(TabContainer)

                Me.tpParts.Enabled = False
                Me.tpProductDiagrams.Enabled = False
                Me.tpProductManuals.Enabled = False
                Me.tpProductSpecs.Enabled = False
                Me.tpProductRetailers.Enabled = False
                Me.tpCrossells.Enabled = False


                'Set the ProductID from Query String
                If Request.QueryString("ProductID") <> "" AndAlso IsNumeric(Request.QueryString("ProductID")) AndAlso Request.QueryString("ProductID") <> 0 Then
                    ViewState("ProductID") = Request.QueryString("ProductID")

                    ResetPanels()

                    InitializeForm()

                    CleanUpOldImages() ' Attempt to cleanup obselete files 

                    If ViewState("ProductID") > 0 Then

                        LoadImageProperites() ' Load 
                        LoadProductImage()
                        ToggleForm(FormMode.PrimaryImage)

                    End If

                Else
                    Response.Redirect("ManageProducts.aspx")
                End If
            Catch ex As Exception
                lblError.Text = ex.Message
                pnlError.Visible = True
                UpdatePanelErrors.Update()
            End Try


        End If
    End Sub
    Private Sub ResetPanels()
        'Reset Panels

        pnlError.Visible = False
        pnlReturnMsg.Visible = False
        UpdatePanelErrors.Update()
    End Sub
    Private Sub InitializeForm()
        'Dim bHasProductDiagram As Boolean
        'Dim sProductDiagram As sProductDiagram = Nothing
        'Dim Product As New Product
        Try
            'If ViewState("ProductID") > 0 Then
            '    ViewState("ProductModelNumber") = CType(fvProductHeader.FindControl("lblProductModelNumber"), Label).Text

            '    sProductDiagram.ProductID = ViewState("ProductID")
            '    bHasProductDiagram = Product.GetProductDiagram(sProductDiagram)

            '    If bHasProductDiagram Then
            '        lblCurrentDiagram.Text = "(current diagram: " & sProductDiagram.FileName & ")"

            '        btnDiagram.ImageUrl = ConfigurationManager.AppSettings("ProductDiagramExternalLinkPath") & sProductDiagram.FileName
            '        btnDiagram.Width = Unit.Pixel(550)
            '        btnDiagram.Attributes.Add("onclick", "javascript:window.open('" & btnDiagram.ImageUrl & "', 'newWindow', 'height=500,width=650,scrollbars=YES,resizable=YES,toolbars=no,status=no,menubar=no,location=no');")
            '        btnDiagram.Visible = True
            '    Else
            '        lblCurrentDiagram.Text = "(current diagram:  none)"
            '    End If
            'End If

        Catch ex As Exception
            lblError.Text = ex.Message
            pnlError.Visible = True
            UpdatePanelErrors.Update()
        Finally
            '  Product = Nothing
        End Try
    End Sub

    Private Sub LoadDiagramFromFile()

        Dim Product As New Product
        Dim sProductDiagram As sProductDiagram
        Dim blnsuccess As Boolean = True

        Try

            sProductDiagram.ProductID = ViewState("ProductID")
            sProductDiagram.Title = "Parts Diagram"
            sProductDiagram.Description = "Parts Diagram"
            sProductDiagram.FileName = ViewState("DiagramFileName")
            sProductDiagram.FilePath = ViewState("DiagramPathOnly")
            sProductDiagram.CreatedBy = Page.User.Identity.Name
            sProductDiagram.CreatedByName = "unknown"
            sProductDiagram.DateCreated = Now

            blnsuccess = Product.AddProductDiagram(sProductDiagram)

            If blnsuccess Then
                Me.lblReturnMsg.Text = "Product Diagram successfully added to model."
                Me.pnlReturnMsg.Visible = True
                Me.pnlError.Visible = False
                UpdatePanelErrors.Update()

                'Update label
                'lblCurrentDiagram.Text = "(current diagram: " & ViewState("DiagramFileName") & ")"
            Else
                Me.lblError.Text = "Error Adding Product Diagram.  Make sure that file type is JPG and try again."
                Me.pnlError.Visible = True
                Me.pnlReturnMsg.Visible = False
                UpdatePanelErrors.Update()
            End If

        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
            blnsuccess = False
        Finally
            Product = Nothing
        End Try
    End Sub
    Private Function UploadDiagramFile() As Boolean
        Dim strSaveLocation As String
        Dim strFilename As String
        Dim blnFileExist As Boolean = True

        Dim bSuccess As Boolean
        Try


            'First Upload the Attachment (If Necessary)
            If Not txtDiagramFile.PostedFile Is Nothing And txtDiagramFile.PostedFile.ContentLength > 0 Then

                strSaveLocation = ConfigurationManager.AppSettings("ProductDiagramStorageLocation").ToString

                strFilename = System.IO.Path.GetFileName(txtDiagramFile.PostedFile.FileName)   'Used when you have to rename a file 

                If LCase(strFilename) Like "*jpg*" Or LCase(strFilename) Like "*jpeg*" Then
                    ViewState("DiagramPath") = strSaveLocation & strFilename

                    'Check to see if file with Samename already Exist
                    Do Until blnFileExist = False
                        If System.IO.File.Exists(strSaveLocation & strFilename) Then
                            blnFileExist = True
                            strFilename = String.Format("{0}_{1}_{2}", System.IO.Path.GetFileNameWithoutExtension(strFilename), Left(System.Guid.NewGuid.ToString, 10), System.IO.Path.GetExtension(strFilename))
                        Else
                            blnFileExist = False
                        End If
                    Loop

                    ViewState("DiagramPath") = strSaveLocation & strFilename
                    ViewState("DiagramPathOnly") = strSaveLocation
                    ViewState("DiagramFileName") = strFilename

                    'Upload & Save the File
                    txtDiagramFile.PostedFile.SaveAs(strSaveLocation & strFilename)

                    bSuccess = True
                Else
                    bSuccess = False
                    Me.lblError.Text = "Incorrect file format.  The selected file type must be JPG.  Please check your file and try again."
                    Me.pnlError.Visible = True
                    Me.pnlReturnMsg.Visible = False
                    UpdatePanelErrors.Update()
                End If

            End If
        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
            bSuccess = False
        Finally

        End Try

        Return bSuccess
    End Function

    Protected Sub btnUploadDiagram_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUploadDiagram.Click
        Dim bsuccess As Boolean
        Dim Product As New Product



        Try
            ResetPanels()

            If txtDiagramFile.Value <> "" And txtDiagramFile.Value <> "test" Then
                bsuccess = Me.UploadDiagramFile()
            Else
                Me.lblError.Text = "No file selected"
                Me.pnlError.Visible = True
                Me.pnlReturnMsg.Visible = False
                UpdatePanelErrors.Update()
            End If

            If bsuccess Then
                Product.DeleteAllProductDiagrams(ViewState("ProductID"))
                LoadDiagramFromFile()
                Me.fvDiagram.DataBind()

                Me.UpdatePanelDiagram.Update()
                Me.UpdatePanelDiagramMain.Update()

            End If



        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally
            Product = Nothing
        End Try
    End Sub

    Protected Sub btnUploadBOM_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUploadBOM.Click

        Dim bsuccess As Boolean

        Try
            ResetPanels()

            If txtBOMFile.Value <> "" And txtBOMFile.Value <> "test" Then
                bsuccess = Me.UploadBOMFile()

            End If

            If bsuccess Then
                LoadBOMFromFile()
                'viewstate("BOMPath") 
            End If

        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally

        End Try
    End Sub

    Private Function DeleteAllProductParts(ByVal ProductID As Integer) As Boolean

        'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed
        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object
        Dim SQLCommand As New SqlCommand   'SQLCommand Object
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Try 'Only one "Try" statement 

            SQLConn.Open() 'Open Database

            'Set the Basic Command Information 
            SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
            SQLCommand.Connection = SQLConn                          'Set the Connection

            SQLCommand.CommandText = "LLFBPS..[DeleteALLProductParts]"


            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ProductID))

            'Set the DataReader
            SQLCommand.ExecuteReader(CommandBehavior.SingleResult)

            'Cleanup
            SQLConn.Close()

        Catch SQLErr As SqlException
            bSuccess = False

        Catch Err As Exception
            bSuccess = False

        Finally
            'Confirm that The SQLDB Connection is closed
            If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
            SQLCommand = Nothing
            SQLConn = Nothing

        End Try

        Return bSuccess

    End Function
    Private Function DeleteAllProductSPECS(ByVal ProductID As Integer) As Boolean

        'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed
        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object
        Dim SQLCommand As New SqlCommand   'SQLCommand Object
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Try 'Only one "Try" statement 

            SQLConn.Open() 'Open Database

            'Set the Basic Command Information 
            SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
            SQLCommand.Connection = SQLConn                          'Set the Connection

            SQLCommand.CommandText = "LLFBPS..[DeleteALLProductSpecifications]"


            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ProductID))

            'Set the DataReader
            SQLCommand.ExecuteReader(CommandBehavior.SingleResult)

            'Cleanup
            SQLConn.Close()

        Catch SQLErr As SqlException
            bSuccess = False

        Catch Err As Exception
            bSuccess = False

        Finally
            'Confirm that The SQLDB Connection is closed
            If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
            SQLCommand = Nothing
            SQLConn = Nothing

        End Try

        Return bSuccess

    End Function
    Private Sub LoadBOMFromFile()

        Dim DS As New System.Data.DataSet
        Dim DT As New DataTable("LFParts")
        Dim DR As DataRow
        Dim DC As DataColumn

        Dim objPart As New Part
        Dim objProduct As New Product
        Dim sPart As sPart = Nothing
        Dim sProduct As sProduct = Nothing

        Dim II As Integer = 0

        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection

        Dim blnsuccess As Boolean = False

        Try
            DC = New DataColumn
            DC.AllowDBNull = True
            DC.DataType = System.Type.GetType("System.String")
            DC.ColumnName = "Key"
            DT.Columns.Add(DC)

            DC = New DataColumn
            DC.AllowDBNull = True
            DC.DataType = System.Type.GetType("System.String")
            DC.ColumnName = "Qty"
            DT.Columns.Add(DC)


            DC = New DataColumn
            DC.AllowDBNull = True
            DC.DataType = System.Type.GetType("System.String")
            DC.ColumnName = "Description"
            DT.Columns.Add(DC)

            DC = New DataColumn
            DC.AllowDBNull = True
            DC.DataType = System.Type.GetType("System.String")
            DC.ColumnName = "PartNumber"
            DT.Columns.Add(DC)

            DC = New DataColumn
            DC.AllowDBNull = True
            DC.DataType = System.Type.GetType("System.String")
            DC.ColumnName = "WarrantyDuration"
            DT.Columns.Add(DC)

            Dim MyFile As System.IO.FileInfo

            Dim I As Integer = 0

            MyFile = New FileInfo(ViewState("BOMPath"))
            If InStr(MyFile.Name, ".xls") Then


                sProduct.ProductID = ViewState("ProductID")
                If objProduct.GetProduct(sProduct) = True Then


                    Try

                        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                              "data source=" & MyFile.FullName & "; Extended Properties=Excel 8.0;")



                        MyCommand = New System.Data.OleDb.OleDbDataAdapter( _
                        "select * from [Sheet1$]", MyConnection)


                        DS = New System.Data.DataSet



                        MyCommand.Fill(DS)
                        MyConnection.Close()


                        DT = DS.Tables(0)

                        Dim CurrentSeriesID As Integer = -1

                        If DT.Rows.Count > 0 Then
                            DeleteAllProductParts(ViewState("ProductID"))


                            For Each DR In DT.Rows

                                Try
                                    If Len(DR(3)) > 0 Then

                                        With sPart
                                            .Active = True
                                            .CreatedBy = Page.User.Identity.Name
                                            .CreatedByName = "unknown"
                                            .DateCreated = Now()
                                            .Description = DR(2)
                                            .Name = DR(2)
                                            .PartNumber = DR(3)
                                            If Not IsDBNull(DR(0)) Then
                                                .CallLetter = Replace(DR(0), "'", "")
                                            Else
                                                .CallLetter = ""
                                            End If

                                            If Not IsDBNull(DR(1)) Then
                                                .ECommerce_Quantity = DR(1)
                                            End If

                                            'If IsNumeric(DR(4)) Then
                                            '    .BaseWarrantyDuration = DR(4)
                                            '    .PartProductWarrantyDuration = DR(4)
                                            'Else
                                            '    'default to 12 if none listed
                                            .BaseWarrantyDuration = 12
                                            .PartProductWarrantyDuration = 12

                                            ' End If
                                            .ProductID = sProduct.ProductID
                                        End With

                                        'Amy bypass me
                                        If objPart.AddPartToProduct(sPart) Then
                                            II += 1
                                            blnsuccess = True
                                        Else
                                            blnsuccess = False
                                        End If

                                    Else 'Bad Row in Spreadsheet

                                        I += 1

                                    End If



                                Catch ex As Exception
                                    'Bad Row in Spreadsheet
                                    I += 1
                                End Try
                            Next

                            If I > 0 Then
                                Me.lblReturnMsg.Text = II & " Parts(s) were successfully added for this model.  " & I & " Part(s) were ignored!"
                                Me.pnlReturnMsg.Visible = True
                                Me.pnlError.Visible = False
                                UpdatePanelErrors.Update()

                            Else
                                Me.lblReturnMsg.Text = II & " Parts(s) were successfully added for this model."
                                Me.pnlReturnMsg.Visible = True
                                Me.pnlError.Visible = False
                                UpdatePanelErrors.Update()
                            End If



                        Else
                            Me.lblError.Text = "There are no records in this table.  Make sure that the first tab in the spreadsheet is named 'Sheet1' and try again"
                            Me.pnlError.Visible = True
                            Me.pnlReturnMsg.Visible = False
                            UpdatePanelErrors.Update()

                        End If

                    Catch exx As Exception
                        Select Case Err.Number
                            Case 5

                                Me.lblError.Text = "No valid records were found.  Make sure that the first tab in the spreadsheet is named 'Sheet1' and try again"
                                Me.pnlError.Visible = True
                                Me.pnlReturnMsg.Visible = False
                                UpdatePanelErrors.Update()
                            Case Else
                                Me.lblError.Text = exx.Message
                                Me.pnlError.Visible = True
                                Me.pnlReturnMsg.Visible = False
                                UpdatePanelErrors.Update()
                        End Select

                    Finally
                        MyFile.Delete()

                        'Rebind data?
                        Me.gvParts.DataBind()
                    End Try

                Else 'Model not found
                End If
            End If
        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally
            objPart = Nothing
            objProduct = Nothing
        End Try
    End Sub
    Private Function UploadBOMFile() As Boolean
        Dim strSaveLocation As String
        Dim strFilename As String
        Dim blnFileExist As Boolean = True

        Dim bSuccess As Boolean
        Try

            'First Upload the Attachment (If Necessary)
            If Not Me.txtBOMFile.PostedFile Is Nothing And Me.txtBOMFile.PostedFile.ContentLength > 0 Then

                strSaveLocation = ConfigurationManager.AppSettings("BOMUploadLocation").ToString

                strFilename = System.IO.Path.GetFileName(Me.txtBOMFile.PostedFile.FileName)   'Used when you have to rename a file 

                If LCase(strFilename) Like "*xls*" Then
                    ViewState("BOMPath") = strSaveLocation & strFilename

                    'Check to see if file with Samename already Exist
                    Do Until blnFileExist = False
                        If System.IO.File.Exists(strSaveLocation & strFilename) Then
                            blnFileExist = True
                            strFilename = String.Format("{0}_{1}_{2}", System.IO.Path.GetFileNameWithoutExtension(strFilename), Left(System.Guid.NewGuid.ToString, 10), System.IO.Path.GetExtension(strFilename))
                        Else
                            blnFileExist = False
                        End If
                    Loop

                    ViewState("BOMPath") = strSaveLocation & strFilename

                    'Upload & Save the File
                    Me.txtBOMFile.PostedFile.SaveAs(strSaveLocation & strFilename)

                    bSuccess = True
                Else
                    bSuccess = False

                    Me.lblError.Text = "Incorrect file format.  The selected file must be an Excel Spreadsheet.  Please check your file and try again."
                    Me.pnlError.Visible = True
                    Me.pnlReturnMsg.Visible = False
                    UpdatePanelErrors.Update()
                End If

            End If
        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
            bSuccess = False
        Finally

        End Try

        Return bSuccess
    End Function
    Private Function UploadSPECFile() As Boolean
        Dim strSaveLocation As String
        Dim strFilename As String
        Dim blnFileExist As Boolean = True

        Dim bSuccess As Boolean
        Try

            'First Upload the Attachment (If Necessary)
            If Not Me.txtSPECSFile.PostedFile Is Nothing And Me.txtSPECSFile.PostedFile.ContentLength > 0 Then

                strSaveLocation = ConfigurationManager.AppSettings("SPECSUploadLocation").ToString

                strFilename = System.IO.Path.GetFileName(Me.txtSPECSFile.PostedFile.FileName)   'Used when you have to rename a file 

                If LCase(strFilename) Like "*xls*" Then
                    ViewState("SPECPath") = strSaveLocation & strFilename

                    'Check to see if file with Samename already Exist
                    Do Until blnFileExist = False
                        If System.IO.File.Exists(strSaveLocation & strFilename) Then
                            blnFileExist = True
                            strFilename = String.Format("{0}_{1}_{2}", System.IO.Path.GetFileNameWithoutExtension(strFilename), Left(System.Guid.NewGuid.ToString, 10), System.IO.Path.GetExtension(strFilename))
                        Else
                            blnFileExist = False
                        End If
                    Loop

                    ViewState("SPECPath") = strSaveLocation & strFilename

                    'Upload & Save the File
                    Me.txtSPECSFile.PostedFile.SaveAs(strSaveLocation & strFilename)

                    bSuccess = True
                Else
                    bSuccess = False

                    Me.lblError.Text = "Incorrect file format.  The selected file must be an Excel Spreadsheet.  Please check your file and try again."
                    Me.pnlError.Visible = True
                    Me.pnlReturnMsg.Visible = False
                    UpdatePanelErrors.Update()
                End If

            End If
        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
            bSuccess = False
        Finally

        End Try

        Return bSuccess
    End Function

    Private Function UploadManualFile(Optional ByVal KnowledgeID As Integer = 0) As Boolean
        Dim blnHasAttachment As Boolean
        Dim strSaveLocation As String
        Dim strFilename As String
        Dim blnFileExist As Boolean = True
        Dim sKnowledgeDocument As New sKnowledgeDocument
        Dim Knowledge As New Knowledge
        Dim bSuccess As Boolean
        Try

            'First Upload the Attachment (If Necessary)
            If Not Me.txtManualFile.PostedFile Is Nothing And Me.txtManualFile.PostedFile.ContentLength > 0 Then
                blnHasAttachment = True
                'strSaveLocation = ConfigurationSettings.AppSettings("FileStorageLocation")  'Used to Save a uploaded File
                strSaveLocation = BuildKnowledgeUploadPath(ConfigurationManager.AppSettings("CompanyID").ToString, ConfigurationManager.AppSettings("BrandID").ToString)

                strFilename = System.IO.Path.GetFileName(Me.txtManualFile.PostedFile.FileName)   'Used when you have to rename a file 

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
                    Me.txtManualFile.PostedFile.SaveAs(strSaveLocation & strFilename)

                    'Add the Image data to the database
                    If AddProductManual(strSaveLocation & strFilename) Then
                        Me.lblReturnMsg.Text = "File Uploaded successfully"
                        Me.pnlReturnMsg.Visible = True
                        Me.pnlError.Visible = False
                        UpdatePanelErrors.Update()
                        Me.gvManuals.DataBind()

                    Else
                        System.IO.File.Delete(sKnowledgeDocument.LocalSystemPath)
                        Me.lblError.Text = "Error uploading file"
                        Me.pnlError.Visible = True
                        UpdatePanelErrors.Update()
                        Me.pnlReturnMsg.Visible = False

                    End If
                Else
                    bSuccess = False

                    Me.lblError.Text = "Incorrect file format.  The selected file must be a PDF file.  Please check your file and try again."
                    Me.pnlError.Visible = True
                    Me.pnlReturnMsg.Visible = False
                    UpdatePanelErrors.Update()
                End If
            Else
                bSuccess = False

                Me.lblError.Text = "No File Selected.  Please select a file and try again."
                Me.pnlError.Visible = True
                Me.pnlReturnMsg.Visible = False
                UpdatePanelErrors.Update()
            End If

        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally
            Knowledge = Nothing
        End Try
    End Function
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

            SQLCommand.CommandText = "LLFBPS..[GetResourceNamesForImageUploadPath]"

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

    Public Function AddProductManual(ByVal LocalSystemPath As String) As Boolean

        Dim bSuccess As Boolean 'Return variable 

        'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

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
            SQLCommand.CommandText = "LLFBPS..[AddProductManual]"     'Stored Procedure Name

            'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Title", System.Data.SqlDbType.VarChar, 250))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 150))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@LocalSystemPath", System.Data.SqlDbType.VarChar, 500))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))

            SQLCommand.Parameters("@Title").Value = Me.txtProductManualTitle.Text

            If IsNumeric(Page.User.Identity.Name) Then
                SQLCommand.Parameters("@CreatedBy").Value = Page.User.Identity.Name
            Else
                SQLCommand.Parameters("@CreatedBy").Value = 0
            End If

            SQLCommand.Parameters("@CreatedName").Value = ""
            SQLCommand.Parameters("@LocalSystemPath").Value = LocalSystemPath
            SQLCommand.Parameters("@ProductID").Value = ViewState("ProductID")

            SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

            If Not SQLCommand.Parameters("@KnowledgeID").Value > 0 Then
                bSuccess = False
            Else
                bSuccess = True
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

    Protected Sub btnUploadManual_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUploadManual.Click
        Dim bsuccess As Boolean

        Try
            ResetPanels()

            If txtManualFile.Value <> "" And txtManualFile.Value <> "test" Then
                bsuccess = Me.UploadManualFile()
            Else
                bsuccess = False

                Me.lblError.Text = "No File Selected.  Please select a file and try again."
                Me.pnlError.Visible = True
                Me.pnlReturnMsg.Visible = False
                UpdatePanelErrors.Update()
            End If

        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally

        End Try
    End Sub


    Private Sub LoadImageProperites()
        Me.lblLargeImageSizeHeight.Text = ConfigurationManager.AppSettings("MaxLargeHeight")
        Me.lblLargeImageSizeWidth.Text = ConfigurationManager.AppSettings("MaxLargeWidth")

        Me.lblSmallImageSizeHeight.Text = ConfigurationManager.AppSettings("MaxSmallHeight")
        Me.lblSmallImageSizeWidth.Text = ConfigurationManager.AppSettings("MaxSmallWidth")

        Me.lblRegularImageSizeHeight.Text = ConfigurationManager.AppSettings("MaxRegularHeight")
        Me.lblRegularImageSizeWidth.Text = ConfigurationManager.AppSettings("MaxRegularWidth")

        Me.lblFeatureImageSizeHeight.Text = ConfigurationManager.AppSettings("MaxFeatureHeight")
        Me.lblFeatureImageSizeWidth.Text = ConfigurationManager.AppSettings("MaxFeatureWidth")

        Me.lblAlternateRegularSizeWidth.Text = ConfigurationManager.AppSettings("MaxAlternateWidth")
        Me.lblAlternateRegularSizeHeight.Text = ConfigurationManager.AppSettings("MaxAlternateHeight")

        Me.lblAlternateSmallSizeWidth.Text = ConfigurationManager.AppSettings("MaxAlternateSmallWidth")
        Me.lblAlternateSmallSizeHeight.Text = ConfigurationManager.AppSettings("MaxAlternateSmallHeight")

        Me.lblConstrain.Text = ConfigurationManager.AppSettings("ConstrainImageonResize")

    End Sub

    Public Sub CleanUpOldImages()
        'Clear this directory on page load. 
        Dim oTempDir As New DirectoryInfo(ConfigurationManager.AppSettings("TempResizeUploadDirectory").ToString)
        Dim oFile As FileInfo

        For Each oFile In oTempDir.GetFiles
            Try
                oFile.Delete()
            Catch ex As Exception

            End Try
        Next
        oTempDir = Nothing
        oFile = Nothing
    End Sub


    Private Sub ToggleForm(ByVal Mode As FormMode)
        Me.pPrimary.Visible = False
        Me.pAlternate.Visible = False

        Select Case Mode
            Case FormMode.PrimaryImage
                Me.lnkAlternate.CssClass = ""
                Me.lnkPrimary.CssClass = "itemselected"
                Me.pPrimary.Visible = True
                'Me.Master.SetDisplayMessage("Welcome to Product Image Management", MessageType.GeneralMessage)

            Case FormMode.AlternateImages
                Me.lnkAlternate.CssClass = "itemselected"
                Me.lnkPrimary.CssClass = ""
                Me.pAlternate.Visible = True
                Me.gvAlternateImages.DataBind()

                'Me.Master.SetDisplayMessage("Welcome to Alternate Product Image Management", MessageType.GeneralMessage)

        End Select
    End Sub

    Private Function GetCurrentImageID(ByVal ImageSize As MyImageSize) As Integer
        Select Case ImageSize
            Case MyImageSize.feature
                Return ViewState("ImageIDFeature")
            Case MyImageSize.large
                Return ViewState("ImageIDLarge")
            Case MyImageSize.regular
                Return ViewState("ImageIDRegular")
            Case MyImageSize.small
                Return ViewState("ImageIDSmall")
            Case Else
                Return 0

        End Select
    End Function

#Region "Database Support Calls"
    Public Function ConvertImagePath(ByVal ImagePath As String) As String

        Try
            Return Utilities.GetImageURL(ImagePath)

        Catch ex As Exception
            'Return ManagerUtilities.GetImagePath(0)
        Finally

        End Try
    End Function

    Private Function UpdateProductSeriesImages(ByVal NewImageID As Integer, ByVal ImageType As MyImageSize) As Boolean
        Dim oProduct As New Product
        Dim sProduct As sProduct = Nothing
        Dim PreviousImageID As Integer = 0
        Dim PreviousImagePath As String = ""

        Dim bSuccess As Boolean = False

        Try
            'Get Current Information
            sProduct.ProductID = ViewState("ProductID")
            sProduct.CompanyID = ConfigurationManager.AppSettings("CompanyID")



            If oProduct.GetProduct(sProduct) Then
                'Update ImageID and save
                Select Case ImageType
                    Case MyImageSize.large
                        PreviousImageID = sProduct.PrimaryResourceImageIDLarge
                        PreviousImagePath = sProduct.ImagePathLarge

                        'Assign new Image ID
                        sProduct.PrimaryResourceImageIDLarge = NewImageID

                    Case MyImageSize.regular
                        PreviousImageID = sProduct.PrimaryResourceImageID
                        PreviousImagePath = sProduct.ImagePath

                        'Assign new Image ID
                        sProduct.PrimaryResourceImageID = NewImageID
                    Case MyImageSize.small

                        PreviousImageID = sProduct.PrimaryResourceImageIDSmall
                        PreviousImagePath = sProduct.ImagePathSmall

                        'Assign new Image ID
                        sProduct.PrimaryResourceImageIDSmall = NewImageID
                    Case MyImageSize.feature

                        PreviousImageID = sProduct.PrimaryResourceImageIDFeature
                        PreviousImagePath = sProduct.ImagePathFeature

                        'Assign new Image ID
                        sProduct.PrimaryResourceImageIDFeature = NewImageID

                    Case MyImageSize.Alternate
                        'Association with Image Resource Junc Table 
                        AddImageFiletoResource(NewImageID)
                End Select


                If oProduct.UpdateProductImages(sProduct) Then
                    'Image Management 
                    'Delete the previous Image from Server 
                    If UCase(ConfigurationManager.AppSettings("Application_CleanupOldImageonInsert")) = "ON" Then
                        If PreviousImageID > 0 Then
                            'Remove Old Image ID and File from Database and Server 
                            CleanUpFile(PreviousImagePath, PreviousImageID)
                        End If
                    End If
                Else
                    CleanUpFile("", PreviousImageID)
                    Throw New Exception("There was a updating Product database table" & oProduct.ErrorMessage)
                End If
            Else
                Throw New Exception("Product could not be located for update." & oProduct.ErrorMessage)
            End If


            bSuccess = True

        Catch ex As Exception
            bSuccess = False
            'Me.Master.SetDisplayMessage("Error Updating Product Image Information: " & ex.Message.ToString, MessageType.ErrorMessage)
            Me.lblError.Text = "Error Updating Product Image Information: " & ex.Message.ToString
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally
            oProduct = Nothing
        End Try

        Return bSuccess
    End Function

    Private Function AddImageFile(ByVal ImageName As String, ByVal ImageDescription As String, ByVal FullImagePath As String) As Integer
        Dim oImage As New BPS_BL.BPS.Image
        Dim sImage As sImage = Nothing
        Try

            sImage.Description = ImageDescription
            sImage.ImagePath = FullImagePath

            If oImage.AddImage(sImage) Then
                AddImageFile = sImage.ImageID
                'Me.Master.SetDisplayMessage("Image " & NewFileName & " Updated Successfully ", MessageType.ConfirmationMessage)

            Else
                Throw New Exception("Error Adding/Updating Image: " & oImage.ErrorMessage)
            End If

        Catch ex As Exception
            AddImageFile = 0
            'Me.Master.SetDisplayMessage("Error Adding/Updating Image: " & ex.Message.ToString, MessageType.ErrorMessage)
            Me.lblError.Text = "Error Adding/Updating Image: " & ex.Message.ToString
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally
            oImage = Nothing
        End Try

    End Function

    Private Function AssociateImage(ByVal ImageID As Integer) As Boolean
        Dim oImage As New BPS_BL.BPS.Image
        Dim sImage As BPS_BL.BPS.sImage = Nothing

        Try
            oImage.AddAssociatedImage(sImage)

        Catch ex As Exception

        Finally
            oImage = Nothing
        End Try
    End Function

    Private Function AddImageFiletoResource(ByVal ImageID As Integer) As Boolean
        Dim oImage As New BPS_BL.BPS.Image
        Try

            If oImage.AddImageToResource(BPS_BL.BPS.ResourceType.Product, ViewState("ProductID"), ImageID) Then

                AddImageFiletoResource = True
                'Me.Master.SetDisplayMessage("Image " & NewFileName & " Association Updated Successfully ", MessageType.ConfirmationMessage)
                Me.gvAlternateImages.DataBind()

            Else
                Throw New Exception("Error Adding/Updating Image Association: " & oImage.ErrorMessage)
            End If

        Catch ex As Exception
            AddImageFiletoResource = False
            'Me.Master.SetDisplayMessage("Error Adding/Updating Image Association: " & ex.Message.ToString, MessageType.ErrorMessage)
            Me.lblError.Text = "Error Adding/Updating Image Association: " & ex.Message.ToString
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()

        Finally
            oImage = Nothing
        End Try

    End Function

    Private Sub LoadProductImage()
        Dim oImage As New BPS_BL.BPS.Image

        Dim sImage As sImage
        sImage.ImageID = 0

        Dim oProduct As New Product
        Dim sProduct As sProduct = Nothing

        sProduct.ProductID = ViewState("ProductID")
        sProduct.CompanyID = ConfigurationManager.AppSettings("CompanyID")
        sProduct.PriceListID = 0 'use default pricelist

        ResourceTypeID = BPS_BL.BPS.ResourceType.Product

        'Get Upload Image Path to be used for adding or updating images
        ' Response.Write(" Set Upload Image Path " & ImagePath)


        'SEt temporary upload file location 
        ViewState("DefaultTempUploadPath") = ConfigurationManager.AppSettings("TempResizeUploadDirectory").ToString



        Try
            If oProduct.GetProduct(sProduct) Then
                'Set page title 
                ImagePath = Utilities.BuildImageUploadPath(sProduct.CompanyID, sProduct.BrandID, ResourceTypeID)

                'Set Production Image File Location 
                ViewState("DefaultUploadPath") = ImagePath
                ViewState("BrandID") = sProduct.BrandID

                ViewState("ProductModelNumber") = sProduct.ProductModelNumber
                'Me.Master.SetDisplayMessage("Image Management for Product " & sProduct.Name, MessageType.GeneralMessage)
                'Me.Master.PageHeader = sProduct.Name & " Images <BR>(" & sProduct.ProductSeriesNumber & ")"
                'Me.Master.PageTitle = "Image Management"

                'Set the Product Series Number

                'Check for current primary image 
                If sProduct.PrimaryResourceImageID > 0 Then

                    Me.imgPrimary.ImageUrl = Utilities.GetImageURL(sProduct.ImagePath)
                    Me.imgRegular.ImageUrl = Utilities.GetImageURL(sProduct.ImagePath)
                    'Me.Master.SetDisplayMessage("To change the primary image for this product, select the ""browse"" button below", Management.MessageType.instructionMessage)
                Else
                    'Me.Master.SetDisplayMessage("To add a primary image, select the ""browse"" button below", Management.MessageType.instructionMessage)
                    Me.imgPrimary.ImageUrl = Utilities.GetImageURL(sProduct.ImagePath) 'Should be default no image 
                    Me.imgRegular.ImageUrl = Utilities.GetImageURL(sProduct.ImagePath)
                End If

                ViewState("ImageIDRegular") = sProduct.PrimaryResourceImageID


                'Check for current Small image 
                If sProduct.PrimaryResourceImageIDSmall > 0 Then
                    'Me.Master.SetDisplayMessage("To change the primary image for this product, select the ""browse"" button below", Management.MessageType.instructionMessage)
                Else
                    'Me.Master.SetDisplayMessage("To add a primary image, select the ""browse"" button below", Management.MessageType.instructionMessage)

                End If
                Me.imgSmall.ImageUrl = Utilities.GetImageURL(sProduct.ImagePathSmall)

                'Store the current ImageIDs 
                ViewState("ImageIDSmall") = sProduct.PrimaryResourceImageIDSmall

                'Check for current Large image 
                If sProduct.PrimaryResourceImageIDLarge > 0 Then
                    Me.imgLarge.ImageUrl = Utilities.GetImageURL(sProduct.ImagePathLarge)
                    'Only allow regenerate if Large...must use that file. 
                    ViewState("ImagePathLarge") = sProduct.ImagePathLarge
                    Me.btnRegenerateImage.Visible = True


                    'Me.Master.SetDisplayMessage("To change the primary image for this product, select the ""browse"" button below", Management.MessageType.instructionMessage)
                Else
                    'Me.Master.SetDisplayMessage("To add a primary image, select the ""browse"" button below", Management.MessageType.instructionMessage)
                    Me.imgLarge.ImageUrl = Utilities.GetImageURL(sProduct.ImagePathLarge)
                    'Me.imgLarge.ImageUrl = Utilities.GetImageURL(ConfigurationManager.AppSettings("NoImageURL"))
                End If

                'Store the current ImageIDs 
                ViewState("ImageIDLarge") = sProduct.PrimaryResourceImageIDLarge
                ViewState("ImageIDLargePath") = sProduct.ImagePathLarge

                'Check for current feature image 
                If sProduct.PrimaryResourceImageIDFeature > 0 Then
                    'Me.Master.SetDisplayMessage("To change the primary image for this product, select the ""browse"" button below", Management.MessageType.instructionMessage)
                Else
                    'Me.Master.SetDisplayMessage("To add a primary image, select the ""browse"" button below", Management.MessageType.instructionMessage)

                End If
                Me.imgFeature.ImageUrl = Utilities.GetImageURL(sProduct.ImagePathFeature)

                'Store the current ImageIDs 
                ViewState("ImageIDFeature") = sProduct.PrimaryResourceImageIDFeature

            Else
                Me.btnRegenerateImage.Visible = False
                Me.btnUpdateImage.Visible = False
                Me.chkAutoResize.Enabled = False

                Me.Master.SetDisplayMessage("Problem loading product" & oProduct.ErrorMessage, Management.MessageType.ErrorMessage)

                'Me.lblError.Text = "This current Product is not available"
                Me.pnlError.Visible = True
                Me.pnlReturnMsg.Visible = False
                UpdatePanelErrors.Update()
                Me.imgPrimary.ImageUrl = Utilities.GetImageURL(ConfigurationManager.AppSettings("NoImageURL"))

            End If

        Catch ex As Exception
            'Me.Master.SetDisplayMessage("Error Loading Image File" & ex.Message, MessageType.ErrorMessage)

            Me.lblError.Text = "Error Loading Image File" & ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally
            oImage = Nothing
            oProduct = Nothing

        End Try


    End Sub

#End Region

#Region "Button Click Events "
    Protected Sub lnkPrimary_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkPrimary.Click
        Me.ToggleForm(FormMode.PrimaryImage)
    End Sub

    Protected Sub lnkAlternate_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkAlternate.Click
        Me.ToggleForm(FormMode.AlternateImages)
    End Sub

    Protected Sub btnUpdateImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdateImage.Click
        Dim NewTempFileName As String = ""


        ResetPanels()
        Dim bSuccess As Boolean = True

        'Upload largest size 
        If FileProductImageUpload.HasFile Then

            If UploadFile(Me.FileProductImageUpload, MyImageSize.baseimage, True) >= 0 Then
                'store the original image path 
                'ViewState("TempFile") = ViewState("DefaultFullUploadPath")
                ' Response.Write("Uploaded file =" & ViewState("TempFile") & "<BR>")

                'store the original file name
                ViewState("TempFile") = FileProductImageUpload.FileName

                'Response.Write("Attemping to resize to Large Image size<BR>")
                'Create Large Image
                If Not Me.CreateResizedImage(ViewState("TempFile"), ViewState("DefaultFullUploadPath"), MyImageSize.large) Then
                    bSuccess = False
                End If

                'Response.Write("Attemping to resize to regular size<BR>")
                'Create regular image
                If Not Me.CreateResizedImage(ViewState("ProductID"), ViewState("DefaultFullUploadPath"), MyImageSize.regular) Then
                    bSuccess = False
                End If

                'Create Small image
                'Response.Write("Attemping to resize to small size<BR>")
                If Not Me.CreateResizedImage(ViewState("ProductID"), ViewState("DefaultFullUploadPath"), MyImageSize.small) Then
                    bSuccess = False
                End If


                'Create feature image
                'Response.Write("Attemping to resize to small size<BR>")
                If Not Me.CreateResizedImage(ViewState("ProductID"), ViewState("DefaultFullUploadPath"), MyImageSize.feature) Then
                    bSuccess = False
                End If
            End If

            Try
                'TODO clean up base image
                'Note: this will probably failed due to locks held during previous resizing...
                'this director is cleared on page load. 

                File.Delete(ViewState("TempFile"))
            Catch ex As Exception
                'This error is to be expected due to the windows/OS still having handle(lock) on file. 
                '  Me.Master.SetDisplayMessage("Error Cleaning up Temporary Files: " & ViewState("TempFile"), MessageType.ErrorMessage)
                'Me.lblError.Text = "Error Cleaning up Temporary Files: " & ViewState("TempFile") & " - " & ex.Message
                'Me.pnlError.Visible = True
                'Me.pnlReturnMsg.Visible = False
                'UpdatePanelErrors.Update()
            End Try

            If bSuccess Then
                Me.LoadProductImage()
            End If
            'Reload image 

        Else
            Me.lblError.Text = "Please select your Product Image using the ""Browse"" button."
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()


            'Me.Master.SetDisplayMessage("Please select your Product Image using the ""Browse"" button.", MessageType.SyntaxMessage)
        End If


    End Sub



    Protected Sub btnAddAlternateImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddAlternateImage.Click
        Try
            Dim NewImageID As Integer = 0
            NewImageID = Me.UploadFile(Me.FileUploadAlternate, MyImageSize.Alternate)

            If NewImageID > 0 Then

                Dim FileName As String = Me.FileUploadAlternate.FileName
                ' Response.Write("Uploaded file  " & ViewState("DefaultFullUploadPath") & "<BR>")
                ViewState("TempFile") = FileUploadAlternate.FileName
                ResizeAlternateImage(NewImageID, ViewState("DefaultFullUploadPath"), MyImageSize.Alternate)

                ''Get the original file name so "_thumnail" can be append to it resized and saved on disk
                'Dim alternateFileName As String = Utilities.GetFileName(Me.FileUploadAlternate.FileName)
                'alternateFileName = alternateFileName & Utilities.parseProductNumberFromFileName(alternateFileName)

                Me.gvAlternateImages.DataBind()

            End If


        Catch ex As Exception

        End Try
    End Sub

    Protected Sub chkAutoResize_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkAutoResize.CheckedChanged
        If Me.chkAutoResize.Checked Then
            Me.pImageSizeSettings.Visible = True
            Me.pProductImages.Visible = False
            Me.imgPrimary.Visible = True
            'Me.Master.SetDisplayMessage("To generate small and regular product images, specify the image properties below for each image", MessageType.GeneralMessage)
        Else
            'Me.Master.SetDisplayMessage("To change a specific product image size, select image file using the ""Browse"" button and click ""Update Image"" to save.", MessageType.GeneralMessage)
            Me.pImageSizeSettings.Visible = False
            Me.pProductImages.Visible = True
            Me.imgPrimary.Visible = False
        End If
    End Sub

    Protected Sub btnUpdateSmallImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdateSmallImage.Click
        ResizeCurrentImage(Me.FileSmallImageUpload, MyImageSize.small)
    End Sub

    Protected Sub btnUpdateRegularImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdateRegularImage.Click
        ResizeCurrentImage(Me.FileRegularImageUpload, MyImageSize.regular)
    End Sub

    Protected Sub btnUpdateLargeImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdateLargeImage.Click
        ResizeCurrentImage(Me.FileLargeImageUpload, MyImageSize.large)
    End Sub


    Protected Sub dsAlternateImages_Deleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsAlternateImages.Deleted
        If e.AffectedRows >= 1 Then
            File.Delete(ViewState("DeleteImagePath"))
            'Me.Master.SetDisplayMessage("File was deleted successfully", MessageType.ConfirmationMessage)
        Else
            If Not IsNothing(e.Exception) Then
                'Me.Master.SetDisplayMessage("Error deleting file: & " & e.Exception.Message.ToString, MessageType.ConfirmationMessage)
            Else
                'Me.Master.SetDisplayMessage("There was a problem deleting file from web site", MessageType.ConfirmationMessage)
            End If
        End If
    End Sub


    Protected Sub gvAlternateImages_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gvAlternateImages.RowDeleting

        ViewState("DeleteImagePath") = CType(Me.gvAlternateImages.Rows(e.RowIndex).FindControl("lblImagePath"), Label).Text
        Try

        Catch ex As Exception
            ViewState("DeleteImagePath") = ""

        End Try
    End Sub



    Protected Sub dsAlternateImages_Updating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) Handles dsAlternateImages.Updating
        Try
            Me.dsAlternateImages.UpdateParameters("Description").DefaultValue = CType(Me.gvAlternateImages.Rows(Me.gvAlternateImages.EditIndex).FindControl("txtEditDescription"), TextBox).Text
            Me.dsAlternateImages.UpdateParameters("ImageID").DefaultValue = CType(Me.gvAlternateImages.Rows(Me.gvAlternateImages.EditIndex).FindControl("lblEditImageID"), Label).Text
        Catch ex As Exception
            'Me.Master.SetDisplayMessage("Error updating Image: " & ex.Message.ToString, MessageType.ErrorMessage)
            Me.lblError.Text = "Error updating Image: " & ex.Message.ToString
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
            e.Cancel = True

        End Try

    End Sub


    Protected Sub btnRegenerateImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRegenerateImage.Click

        Dim bSuccess As Boolean = False

        'Get the Path to the large Image File...use this file to recreate the other sizes
        Dim TemplateFilePath As String = ""

        Try
            If ViewState("ImageIDLarge") > 0 Then
                TemplateFilePath = ViewState("ImageIDLargePath")
            Else
                'Load default image from web config..this is usually the large size 
                TemplateFilePath = ConfigurationManager.AppSettings("NoImageURL")
            End If

            'Me.CreateResizedImage(viewstate("ProductID"), TemplateFilePath, MyImageSize.large)

            If Not Me.CreateResizedImage(ViewState("ProductID"), TemplateFilePath, MyImageSize.regular) Then
                Throw New Exception("Error Creating Regular Image Size")
            End If

            If Not Me.CreateResizedImage(ViewState("ProductID"), TemplateFilePath, MyImageSize.small) Then
                Throw New Exception("Error Creating Small Image Size")
            End If

            If Not Me.CreateResizedImage(ViewState("ProductID"), TemplateFilePath, MyImageSize.feature) Then
                Throw New Exception("Error Creating Feature Image Size")
            End If

            bSuccess = True

            If bSuccess Then
                Me.LoadProductImage()
            End If

        Catch ex As Exception
            'Me.Master.SetDisplayMessage("There was a problem processing image: " & ex.Message.ToString, MessageType.ErrorMessage)
            Me.lblError.Text = "There was a problem processing image: " & ex.Message.ToString
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        End Try

    End Sub
#End Region

#Region "Image Resizing & Upload Support Functions"

    Public Function ImageResizer(ByVal origImage As System.Drawing.Image, ByVal NewWidth As Integer, ByVal NewHeight As Integer) As Drawing.Image

        Dim thumbnail As New System.Drawing.Bitmap(NewWidth, NewHeight)
        Dim graphic As System.Drawing.Graphics

        graphic = System.Drawing.Graphics.FromImage(thumbnail)
        Dim Info As System.Drawing.Imaging.ImageCodecInfo()
        Info = System.Drawing.Imaging.ImageCodecInfo.GetImageEncoders()
        Dim Params As System.Drawing.Imaging.EncoderParameters
        Params = New System.Drawing.Imaging.EncoderParameters(1)
        Params.Param(0) = New System.Drawing.Imaging.EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 100L)

        'set quality properties
        graphic.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic
        graphic.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality
        graphic.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality
        graphic.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality

        graphic.DrawImage(origImage, 0, 0, NewWidth, NewHeight)

        Return thumbnail


    End Function

    Private Function ResizeAlternateImage(ByVal OriginalImageID As Integer, ByVal ImageUrl As String, ByVal ImageSize As MyImageSize) As Boolean
        Dim thumbNailImg As System.Drawing.Image
        Dim ResizedImg As System.Drawing.Image
        Dim fullSizeImg As System.Drawing.Image


        Try
            fullSizeImg = System.Drawing.Image.FromFile(ImageUrl)

            SetImageProperties(fullSizeImg, ImageSize, ImageUrl)
            ' Response.Write("Image is about to be resized by object<BR>")
            ResizedImg = ImageResizer(fullSizeImg, imageWidth, imageHeight)
            ' Response.Write("Image has been resized..prepare to save to disk<BR>")

            Dim FullPath As String = ImagePath & NewFileName
            'Response.Write("Image save path = to : " & FullPath & "<BR>")
            'Response.Write("Image has been resize...now save to disk <BR>")
            'ResizedImg.Save(FullPath, ImageFormat.Jpeg)

            If InStr(ViewState("TempFile").ToString.ToLower, "png") > 0 Then
                ResizedImg.Save(FullPath, Imaging.ImageFormat.Png)
            Else
                ResizedImg.Save(FullPath, Imaging.ImageFormat.Jpeg)
            End If

            'Response.Write("Image has been saved to : " & FullPath & "<BR>")


            Dim NewResizedImageID As Integer = AddImageFile(NewFileName, "Alternate Image View", ImagePath & NewFileName)

            If NewResizedImageID > 0 Then
                Me.AddImageFiletoResource(NewResizedImageID)
            Else
                Throw New Exception("Error saving resized alternate image")
            End If

            'Generate Thumbnail for image. 

            Dim NewThumbnailFileName As String = Utilities.GetFileName(NewFileName)
            'NewThumbnailFileName = Utilities.parseProductNumberFromFileName(NewThumbnailFileName) & "_thumbnail.jpg"



            If InStr(ViewState("TempFile").ToString.ToLower, "png") > 0 Then
                NewThumbnailFileName = Utilities.parseProductNumberFromFileName(NewThumbnailFileName) & "_thumbnail.png"
            Else
                NewThumbnailFileName = Utilities.parseProductNumberFromFileName(NewThumbnailFileName) & "_thumbnail.jpg"
            End If

        




            SetImageProperties(fullSizeImg, MyImageSize.AlternateThumb, ImageUrl)
            '  Response.Write("Image thumbnail has been resized..prepare to save to disk<BR>")
            thumbNailImg = ImageResizer(fullSizeImg, imageWidth, imageHeight)
            FullPath = ImagePath & NewThumbnailFileName
            '  Response.Write("saved thumbnail to : " & FullPath & "<BR>")

            'thumbNailImg.Save(FullPath, ImageFormat.Jpeg)


            If InStr(ViewState("TempFile").ToString.ToLower, "png") > 0 Then
                thumbNailImg.Save(FullPath, Imaging.ImageFormat.Png)
            Else
                thumbNailImg.Save(FullPath, Imaging.ImageFormat.Jpeg)
            End If


            '  Response.Write("Image has been saved to : " & FullPath & "<BR>")


            CleanUpFile(ImageUrl, OriginalImageID) 'Delete the original file. 
            ' Response.Write("clean up file")

            'Clean up / Dispose...
            thumbNailImg.Dispose()
            ResizedImg.Dispose()


            ResizeAlternateImage = True

        Catch ex As Exception
            ResizeAlternateImage = False
        Finally
            thumbNailImg = Nothing
            fullSizeImg = Nothing
            ResizedImg = Nothing
        End Try

    End Function

    Public Function CreateResizedImage(ByVal ProductSeriesNumber As String, ByVal TemplateImageUrl As String, Optional ByVal ImageSize As MyImageSize = MyImageSize.regular, Optional ByVal StoreInDataBase As Boolean = True) As Boolean

        Dim bSuccess As Boolean = False
        Dim fullImagePath As String = ""
        Dim fullSizeImg As System.Drawing.Image
        Dim thumbNailImg As System.Drawing.Image

        fullSizeImg = System.Drawing.Image.FromFile(TemplateImageUrl)
        SetImageProperties(fullSizeImg, ImageSize, TemplateImageUrl) 'Determine the size of the image by size

        Try
            If imageHeight > 0 And imageWidth > 0 Then
                ' Response.Write("Resize image to height: " & imageHeight & " width: " & imageWidth & "<BR>")
                thumbNailImg = ImageResizer(fullSizeImg, imageWidth, imageHeight)

                fullImagePath = ImagePath & NewFileName
                'Response.Write("About to save new file =" & fullImagePath & "<BR>")
                'thumbNailImg.Save(fullImagePath, ImageFormat.Jpeg)
                If InStr(ViewState("TempFile").ToString.ToLower, "png") > 0 Then
                    thumbNailImg.Save(fullImagePath, Imaging.ImageFormat.Png)
                Else
                    thumbNailImg.Save(fullImagePath, Imaging.ImageFormat.Jpeg)
                End If


                'Response.Write("new file was saved=" & fullImagePath & "<BR>")

                'Clean up / Dispose...
                fullSizeImg.Dispose()
                thumbNailImg.Dispose()
                thumbNailImg = Nothing
                fullSizeImg = Nothing


                If StoreInDataBase Then

                    'Add Image to Image Table 
                    Dim NewTempImageID = Me.AddImageFile(NewFileName, NewFileName, fullImagePath)

                    If NewTempImageID > 0 Then
                        If UpdateProductSeriesImages(NewTempImageID, ImageSize) Then

                        Else
                            Throw New Exception("There was an error saving new image information to database...Image resize was aborted: " & NewFileName)
                        End If
                    End If

                End If

            Else
                ' Response.Write("Problem with image sizes ...save a default image to " & ImagePath & NewFileName & " height: " & imageHeight & " width: " & imageWidth & "<BR>")
                'fullSizeImg.Save(Response.OutputStream, ImageFormat.Gif)
                'fullSizeImg.Save(ImagePath & NewFileName, ImageFormat.Jpeg)

                If InStr(ViewState("TempFile").ToString.ToLower, "png") > 0 Then
                    fullSizeImg.Save(ImagePath & NewFileName, Imaging.ImageFormat.Png)
                Else
                    fullSizeImg.Save(ImagePath & NewFileName, Imaging.ImageFormat.Jpeg)
                End If

            End If

            bSuccess = True


        Catch ex As Exception
            bSuccess = False
            'Clean up / Dispose...
            thumbNailImg = Nothing
            fullSizeImg = Nothing
            'Response.Write("Clean up original uploaded file: " & fullImagePath)
            CleanUpFile(fullImagePath)
            'Me.Master.SetDisplayMessage(ex.Message.ToString, MessageType.ErrorMessage)
            Me.lblError.Text = ex.Message.ToString
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally



        End Try

        Return bSuccess

    End Function

    Sub SetImageProperties(ByVal ImageFile As System.Drawing.Image, ByVal ImageSize As MyImageSize, ByVal TemplateImageUrl As String)
        Dim WidthRatio As Double = 0
        Dim HeightRatio As Double = 0
        Dim Ratio As Double = 0
        ImagePath = ""



        Select Case ImageSize
            Case MyImageSize.large
                If Not IsNothing(ImageFile) Then
                    If ImageFile.Height > imageHeight Then

                        If UCase(ConfigurationManager.AppSettings("ConstrainImageonResize")) = "ON" Then
                            HeightRatio = ImageFile.Height / ConfigurationManager.AppSettings("MaxLargeHeight")
                            WidthRatio = ImageFile.Width / ConfigurationManager.AppSettings("MaxLargeWidth")

                            Ratio = Math.Max(WidthRatio, HeightRatio)

                            imageWidth = ImageFile.Width / Ratio
                            imageHeight = ImageFile.Height / Ratio
                        Else
                            imageWidth = ConfigurationManager.AppSettings("MaxLargeWidth")
                            imageHeight = ConfigurationManager.AppSettings("MaxLargeHeight")
                        End If
                    End If
                Else
                    imageWidth = ConfigurationManager.AppSettings("MaxLargeWidth")
                    imageHeight = ConfigurationManager.AppSettings("MaxLargeHeight")
                End If

                ImagePath = ViewState("DefaultUploadPath")
                'NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_large.jpg"


                If InStr(TemplateImageUrl.ToString.ToLower, "png") > 0 Then
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_large.png"
                Else
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_large.jpg"
                End If




            Case MyImageSize.regular
                If Not IsNothing(ImageFile) Then

                    If UCase(ConfigurationManager.AppSettings("ConstrainImageonResize")) = "ON" Then
                        HeightRatio = ImageFile.Height / ConfigurationManager.AppSettings("MaxRegularHeight")
                        WidthRatio = ImageFile.Width / ConfigurationManager.AppSettings("MaxRegularWidth")

                        Ratio = Math.Max(WidthRatio, HeightRatio)
                        imageWidth = ImageFile.Width / Ratio
                        imageHeight = ImageFile.Height / Ratio
                    Else
                        imageWidth = ConfigurationManager.AppSettings("MaxRegularWidth")
                        imageHeight = ConfigurationManager.AppSettings("MaxRegularHeight")
                    End If
                Else
                    imageWidth = ConfigurationManager.AppSettings("MaxRegularWidth")
                    imageHeight = ConfigurationManager.AppSettings("MaxRegularHeight")
                End If
                ImagePath = ViewState("DefaultUploadPath")
                'NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_regular.jpg"

                If InStr(TemplateImageUrl.ToString.ToLower, "png") > 0 Then
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_regular.png"
                Else
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_regular.jpg"
                End If



            Case MyImageSize.small
                If Not IsNothing(ImageFile) Then

                    If UCase(ConfigurationManager.AppSettings("ConstrainImageonResize")) = "ON" Then
                        HeightRatio = ImageFile.Height / ConfigurationManager.AppSettings("MaxSmallHeight")
                        WidthRatio = ImageFile.Width / ConfigurationManager.AppSettings("MaxSmallWidth")

                        Ratio = Math.Max(WidthRatio, HeightRatio)
                        imageWidth = ImageFile.Width / Ratio
                        imageHeight = ImageFile.Height / Ratio
                    Else
                        imageWidth = ConfigurationManager.AppSettings("MaxSmallWidth")
                        imageHeight = ConfigurationManager.AppSettings("MaxSmallHeight")
                    End If
                Else
                    imageWidth = ConfigurationManager.AppSettings("MaxSmallWidth")
                    imageHeight = ConfigurationManager.AppSettings("MaxSmallHeight")
                End If

                ImagePath = ViewState("DefaultUploadPath")
                'NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_small.jpg"

                If InStr(TemplateImageUrl.ToString.ToLower, "png") > 0 Then
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_small.png"
                Else
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_small.jpg"
                End If


            Case MyImageSize.feature
                If Not IsNothing(ImageFile) Then

                    If UCase(ConfigurationManager.AppSettings("ConstrainImageonResize")) = "ON" Then
                        HeightRatio = ImageFile.Height / ConfigurationManager.AppSettings("MaxFeatureHeight")
                        WidthRatio = ImageFile.Width / ConfigurationManager.AppSettings("MaxFeatureWidth")

                        Ratio = Math.Max(WidthRatio, HeightRatio)
                        imageWidth = ImageFile.Width / Ratio
                        imageHeight = ImageFile.Height / Ratio
                    Else
                        imageWidth = ConfigurationManager.AppSettings("MaxFeatureWidth")
                        imageHeight = ConfigurationManager.AppSettings("MaxFeatureHeight")
                    End If
                Else
                    imageWidth = ConfigurationManager.AppSettings("MaxFeatureWidth")
                    imageHeight = ConfigurationManager.AppSettings("MaxFeatureHeight")
                End If

                ImagePath = ViewState("DefaultUploadPath")
                'NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_feature.jpg"

                If InStr(TemplateImageUrl.ToString.ToLower, "png") > 0 Then
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_feature.png"
                Else
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_feature.jpg"
                End If

            Case MyImageSize.Alternate

                If Not IsNothing(ImageFile) Then
                    If UCase(ConfigurationManager.AppSettings("ConstrainImageonResize")) = "ON" Then
                        HeightRatio = ImageFile.Height / ConfigurationManager.AppSettings("MaxAlternateHeight")
                        WidthRatio = ImageFile.Width / ConfigurationManager.AppSettings("MaxAlternateWidth")

                        Ratio = Math.Max(WidthRatio, HeightRatio)
                        imageWidth = ImageFile.Width / Ratio
                        imageHeight = ImageFile.Height / Ratio
                    Else
                        imageWidth = ConfigurationManager.AppSettings("MaxAlternateWidth")
                        imageHeight = ConfigurationManager.AppSettings("MaxAlternateHeight")
                    End If
                Else
                    imageWidth = ConfigurationManager.AppSettings("MaxAlternateWidth")
                    imageHeight = ConfigurationManager.AppSettings("MaxAlternateHeight")
                End If

                ImagePath = ViewState("DefaultUploadPath")
                'NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_alternate_" & Left(System.Guid.NewGuid.ToString, 10) & ".jpg"
                If InStr(ViewState("TempFile").ToString.ToLower, "png") > 0 Then
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_alternate_" & Left(System.Guid.NewGuid.ToString, 10) & ".png"
                Else
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_alternate_" & Left(System.Guid.NewGuid.ToString, 10) & ".jpg"
                End If

            Case MyImageSize.AlternateThumb

                If Not IsNothing(ImageFile) Then
                    If UCase(ConfigurationManager.AppSettings("ConstrainImageonResize")) = "ON" Then
                        HeightRatio = ImageFile.Height / ConfigurationManager.AppSettings("MaxAlternateSmallHeight")
                        WidthRatio = ImageFile.Width / ConfigurationManager.AppSettings("MaxAlternateSmallWidth")

                        Ratio = Math.Max(WidthRatio, HeightRatio)
                        imageWidth = ImageFile.Width / Ratio
                        imageHeight = ImageFile.Height / Ratio
                    Else
                        imageWidth = ConfigurationManager.AppSettings("MaxAlternateSmallWidth")
                        imageHeight = ConfigurationManager.AppSettings("MaxAlternateSmallHeight")
                    End If
                Else
                    imageWidth = ConfigurationManager.AppSettings("MaxAlternateSmallWidth")
                    imageHeight = ConfigurationManager.AppSettings("MaxAlternateSmallHeight")

                End If

                ImagePath = ViewState("DefaultUploadPath")
                'TODO Need to save alternate file name with same guid...+ _thumb
                'NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_alternate_thumb" & Left(System.Guid.NewGuid.ToString, 10) & ".jpg"

                If InStr(TemplateImageUrl.ToString.ToLower, "png") > 0 Then
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_alternate_thumb" & Left(System.Guid.NewGuid.ToString, 10) & ".png"
                Else
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_alternate_thumb" & Left(System.Guid.NewGuid.ToString, 10) & ".jpg"
                End If

            Case MyImageSize.baseimage 'Original Image Upload...doesn't resize
                If Not IsNothing(ImageFile) Then
                    If UCase(ConfigurationManager.AppSettings("ConstrainImageonResize")) = "ON" Then
                        'HeightRatio = ImageFile.Height / ConfigurationManager.AppSettings("MaxLargeHeight")
                        'WidthRatio = ImageFile.Width / ConfigurationManager.AppSettings("MaxLargeWidth")

                        'Ratio = Math.Max(WidthRatio, HeightRatio)
                        'imageWidth = ImageFile.Width / Ratio
                        'imageHeight = ImageFile.Height / Ratio
                    Else
                        ' imageWidth = ConfigurationManager.AppSettings("MaxLargeWidth")
                        ' imageHeight = ConfigurationManager.AppSettings("MaxLargeHeight")
                    End If


                Else
                    'imageWidth = ConfigurationManager.AppSettings("MaxLargeWidth")
                    'imageHeight = ConfigurationManager.AppSettings("MaxLargeHeight")
                End If

                ImagePath = ViewState("DefaultUploadPath")
                'NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_baseimage.jpg"

                If InStr(TemplateImageUrl, "png") > 0 Then
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_baseimage.png"
                Else
                    NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ViewState("BrandID") & "_" & ViewState("ProductModelNumber") & "_baseimage.jpg"
                End If



        End Select
    End Sub


    Private Sub ResizeCurrentImage(ByRef obj As WebControls.FileUpload, ByVal ImageSize As MyImageSize, Optional ByVal ImageID As Integer = 0)

        'Update file to temporary directory
        Dim tempDirectory As String = ConfigurationManager.AppSettings("TempResizeUploadDirectory").ToString
        Dim fullPath As String = ""
        Dim bSuccess As Boolean = False
        Dim ResizedImg As System.Drawing.Image
        Dim fullSizeImg As System.Drawing.Image

        Try

            If obj.HasFile Then
                Dim ClientImage As System.Drawing.Image = Nothing

                'Set image name, imageWidth, imageHeight, path & etc...
                SetImageProperties(ClientImage, ImageSize, NewFileName)
                fullPath = tempDirectory & NewFileName
                obj.PostedFile.SaveAs(fullPath)

            End If

            'Resize file to correct size based on imagesize 
            fullSizeImg = System.Drawing.Image.FromFile(fullPath)

            SetImageProperties(fullSizeImg, ImageSize, NewFileName)

            ResizedImg = ImageResizer(fullSizeImg, imageWidth, imageHeight)

            Dim ProductionFullPath As String = ImagePath & NewFileName

            'overwrite existing file
            'ResizedImg.Save(ProductionFullPath, ImageFormat.Jpeg)
            If InStr(NewFileName.ToString.ToLower, "png") > 0 Then
                ResizedImg.Save(ProductionFullPath, Imaging.ImageFormat.Png)
            Else
                ResizedImg.Save(ProductionFullPath, Imaging.ImageFormat.Jpeg)
            End If


            'Clean up / Dispose...
            fullSizeImg.Dispose()

            ResizedImg.Dispose()

            fullSizeImg = Nothing
            ResizedImg = Nothing

            'delete temporary file 
            Try
                File.Delete(fullPath)
            Catch ex As Exception
                'Me.Master.SetDisplayMessage("Error Removing temp file: " & fullPath & " - " & ex.Message.ToString, MessageType.ErrorMessage)

                Me.lblError.Text = "Error Removing temp file: " & fullPath & " - " & ex.Message.ToString
                Me.pnlError.Visible = True
                Me.pnlReturnMsg.Visible = False
                UpdatePanelErrors.Update()
            End Try

            'Check to see if replace new image or existing image 
            If Not GetCurrentImageID(ImageSize) > 0 Then
                'Add New Image
                Dim NewImageID As Integer = 0

                NewImageID = Me.AddImageFile(NewFileName, NewFileName, ProductionFullPath)
                If Not Me.UpdateProductSeriesImages(NewImageID, ImageSize) Then
                    Throw New Exception("Error Updating Product Image Data")
                End If
            End If

            'Reload Image Data
            Me.LoadProductImage()
        Catch ex As Exception
            'Me.Master.SetDisplayMessage("Error resizing file: " & fullPath & " - " & ex.Message.ToString, MessageType.ErrorMessage)
            Me.lblError.Text = "Error resizing file: " & fullPath & " - " & ex.Message.ToString
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()

        End Try



    End Sub

    Private Function UploadFile(ByRef obj As WebControls.FileUpload, ByVal ImageSize As MyImageSize, Optional ByVal UseTempDir As Boolean = False) As Integer
        Dim bSuccess As Boolean = False
        Dim NewPrimaryImageID As Integer = 0
        Try
            If (obj.HasFile) Then
                Dim ClientImage As System.Drawing.Image = Nothing

                'Set image name, path & etc...
                SetImageProperties(ClientImage, ImageSize, obj.FileName)

                If UseTempDir = True Then
                    ViewState("DefaultFullUploadPath") = ConfigurationManager.AppSettings("TempResizeUploadDirectory").ToString.ToString & NewFileName
                    ViewState("TempFileName") = NewFileName
                Else
                    ViewState("DefaultFullUploadPath") = ViewState("DefaultUploadPath") & NewFileName
                End If

                obj.PostedFile.SaveAs(ViewState("DefaultFullUploadPath"))

                bSuccess = True

                'Create instance of image and check width to ensure dimensions are large ensure to be 
                'resized. 
                ClientImage = System.Drawing.Image.FromFile(ViewState("DefaultFullUploadPath"))

                Dim oFile As New FileInfo(ViewState("DefaultFullUploadPath"))
                'Make sure for furture update that file is not set to read only
                If oFile.IsReadOnly Then
                    File.SetAttributes(ViewState("DefaultFullUploadPath"), FileAttributes.Normal)
                End If
                oFile = Nothing
                'Enforce Large when auto resizing config check is on and enabled. 
                If ImageSize = MyImageSize.large And Me.chkAutoResize.Checked Then
                    If UCase(ConfigurationManager.AppSettings("Application_EnforceMaxWidthOnUpload")) = "ON" Then
                        If ClientImage.Width < ConfigurationManager.AppSettings("MaxLargeWidth") Then
                            Throw New Exception("Primary Product image must be at least " & ConfigurationManager.AppSettings("MaxLargeWidth") & " Pixels in Width. Image " & NewFileName & " upload was aborted")
                        End If
                    End If
                End If

                'Don't store base image into database 
                If ImageSize <> MyImageSize.baseimage Then
                    NewPrimaryImageID = Me.AddImageFile(NewFileName, NewFileName, ViewState("DefaultFullUploadPath"))
                Else
                    NewPrimaryImageID = 0
                End If

                'Assign New Image 
                If NewPrimaryImageID > 0 Then
                    If Me.UpdateProductSeriesImages(NewPrimaryImageID, ImageSize) Then

                    Else
                        Throw New Exception("There was a problem assigning New Product Image , Image upload has been aborted: " & NewFileName)
                    End If
                Else
                    If ImageSize <> MyImageSize.baseimage Then
                        Throw New Exception("There was a problem saving Primary Image to database, Image upload has been aborted: " & NewFileName)
                    End If

                End If

            End If

        Catch ex As Exception
            Me.CleanUpFile(ViewState("DefaultFullUploadPath"), NewPrimaryImageID)
            NewPrimaryImageID = 0 'Return 0 is ImageID was not added. 
            'Me.Master.SetDisplayMessage("Error Uploading File: " & ex.Message.ToString, MessageType.ErrorMessage)
            Me.lblError.Text = "Error Uploading File: " & ex.Message.ToString
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()

        End Try

        Return NewPrimaryImageID
    End Function

    Private Function CleanUpFile(Optional ByVal FilePath As String = "", Optional ByVal ImageID As Integer = 0) As Boolean

        Dim bSuccess As Boolean = False
        Dim oImage As New BPS_BL.BPS.Image

        Try
            If ImageID > 0 Then
                'Delete the Image from the Image Table
                If oImage.DeleteImage(ImageID) Then

                End If
            End If

            ''TODO for testing only...so not affect product data
            'If InStr(FilePath, "\\192.168.225.181\d$") > 0 Then
            '    If FilePath <> "" Then
            '        File.Delete(FilePath)
            '    End If
            'End If


            bSuccess = True
        Catch ex As Exception
            bSuccess = False
        Finally
            oImage = Nothing
        End Try
    End Function
#End Region

#Region " Enumeration "
    Private Enum FormMode
        PrimaryImage = 0
        AlternateImages = 1
    End Enum
    Public Enum MyImageSize
        small = 0
        regular = 1
        large = 2
        Alternate = 3
        AlternateThumb = 6
        baseimage = 4
        feature = 5
    End Enum
#End Region
    'Protected Sub fvProductHeader_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles fvProductHeader.DataBound
    '    'If item is not supported then do not show obsolete panels
    '    If ViewState("ProductID") > 0 AndAlso fvProductHeader.DataItem("IsSupported") = False Then
    '         Me.tpParts.Enabled = False
    '          Me.tpProductDiagrams.Enabled = False
    '           Me.tpProductManuals.Enabled = False
    '    Me.tpProductSpecs.Enabled = False
    '            Me.tpProductRetailers.Enabled = False

    '    End If
    'End Sub

    Protected Sub fvDiagram_DataBound(ByVal sender As Object, ByVal e As System.EventArgs) Handles fvDiagram.DataBound
        If fvDiagram.DataItemCount > 0 Then

            Me.lblCurrentDiagram.Text = "current diagram:  " & fvDiagram.DataItem("FileName")

            CType(fvDiagram.FindControl("btnDiagram"), ImageButton).Attributes.Add("onclick", "javascript:window.open('" & CType(fvDiagram.FindControl("btnDiagram"), ImageButton).ImageUrl & "', 'newWindow', 'height=500,width=650,scrollbars=YES,resizable=YES,toolbars=no,status=no,menubar=no,location=no');")
            CType(fvDiagram.FindControl("btnDiagram"), ImageButton).Visible = True
        Else
            Me.lblCurrentDiagram.Text = "(current diagram:  none)"

        End If
    End Sub


    Protected Sub btnRetailerMoveLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRetailerMoveLeft.Click
        Dim SelectedItem As ListItem
        Try

            For Each SelectedItem In lbProductRetailers.Items
                If SelectedItem.Selected = True Then

                    Me.dsProductRetailers.UpdateParameters("ProductID").DefaultValue = ViewState("ProductID")
                    Me.dsProductRetailers.UpdateParameters("RetailerID").DefaultValue = SelectedItem.Value
                    Me.dsProductRetailers.UpdateParameters("IsProductRetailer").DefaultValue = False

                    dsProductRetailers.Update()

                End If

            Next

            lbProductRetailers.ClearSelection()

            'Refresh Update panel (Panel is set to update conditionally
            Me.UpdatePanelRetailers.Update()

            Me.lblReturnMsg.Text = "Product Retailers updated successfully"
            Me.pnlReturnMsg.Visible = True
            Me.pnlError.Visible = False
            UpdatePanelErrors.Update()

        Catch ex As Exception
        Finally
            SelectedItem = Nothing
        End Try
    End Sub

    Protected Sub btnRetailerMoveRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnRetailerMoveRight.Click

        Dim SelectedItem As ListItem
        Try
            For Each SelectedItem In lbAvailableRetailers.Items
                If SelectedItem.Selected = True Then
                    Me.dsProductRetailers.UpdateParameters("ProductID").DefaultValue = ViewState("ProductID")
                    Me.dsProductRetailers.UpdateParameters("RetailerID").DefaultValue = SelectedItem.Value
                    Me.dsProductRetailers.UpdateParameters("IsProductRetailer").DefaultValue = True

                    dsProductRetailers.Update()

                End If

            Next

            lbAvailableRetailers.ClearSelection()

            'Refresh Update panel (Panel is set to update conditionally
            Me.UpdatePanelRetailers.Update()


            Me.lblReturnMsg.Text = "Product Retailers updated successfully"
            Me.pnlReturnMsg.Visible = True
            Me.pnlError.Visible = False
            UpdatePanelErrors.Update()

        Catch ex As Exception
        Finally
            SelectedItem = Nothing
        End Try

    End Sub


    Protected Sub btnCrossSellMoveLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCrossSellMoveLeft.Click
        Dim SelectedItem As ListItem
        Try

            For Each SelectedItem In lbCrossSells.Items
                If SelectedItem.Selected = True Then
                    Me.dsCrossSells.UpdateParameters("ProductID").DefaultValue = ViewState("ProductID")
                    Me.dsCrossSells.UpdateParameters("AssociatedProductID").DefaultValue = SelectedItem.Value
                    Me.dsCrossSells.UpdateParameters("IsCrossSell").DefaultValue = False

                    dsCrossSells.Update()

                End If

            Next

            lbCSAvailableProducts.ClearSelection()

            'Refresh Update panel (Panel is set to update conditionally
            Me.UpdatePanelCrossells.Update()

            Me.lblReturnMsg.Text = "Cross Sells updated successfully"
            Me.pnlReturnMsg.Visible = True
            Me.pnlError.Visible = False
            UpdatePanelErrors.Update()
        Catch ex As Exception
        Finally
            SelectedItem = Nothing
        End Try

    End Sub

    Protected Sub btnCrossSellMoveRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCrossSellMoveRight.Click
        Dim SelectedItem As ListItem
        Try
            For Each SelectedItem In lbCSAvailableProducts.Items
                If SelectedItem.Selected = True Then
                    Me.dsCrossSells.UpdateParameters("ProductID").DefaultValue = ViewState("ProductID")
                    Me.dsCrossSells.UpdateParameters("AssociatedProductID").DefaultValue = SelectedItem.Value
                    Me.dsCrossSells.UpdateParameters("IsCrossSell").DefaultValue = True

                    dsCrossSells.Update()

                End If

            Next

            lbCSAvailableProducts.ClearSelection()

            'Refresh Update panel (Panel is set to update conditionally
            Me.UpdatePanelCrossells.Update()

            Me.lblReturnMsg.Text = "Cross Sells updated successfully"
            Me.pnlReturnMsg.Visible = True
            Me.pnlError.Visible = False
            UpdatePanelErrors.Update()

        Catch ex As Exception
        Finally
            SelectedItem = Nothing
        End Try

    End Sub

    Protected Sub dsProduct_Updated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsProduct.Updated
        Me.lblReturnMsg.Text = "Product updated successfully"
        Me.pnlReturnMsg.Visible = True
        Me.pnlError.Visible = False
        UpdatePanelErrors.Update()
    End Sub

    Protected Sub dsProductDescription_Updated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsProductDescription.Updated
        Response.Redirect(String.Format("ManageProduct.aspx?ProductID={0}", Request.QueryString("ProductID")))

    End Sub


    Private Sub LoadSPECSFromFile()

        Dim DS As New System.Data.DataSet
        Dim DT As New DataTable("ProductSpecs")
        Dim DR As DataRow
        Dim DC As DataColumn

        Dim objPart As New Part
        Dim objProduct As New Product
        Dim sPart As sPart = Nothing
        Dim sProduct As sProduct = Nothing


        Dim II As Integer = 0

        Dim MyCommand As System.Data.OleDb.OleDbDataAdapter
        Dim MyConnection As System.Data.OleDb.OleDbConnection

        Dim blnsuccess As Boolean = False

        Try
            DC = New DataColumn
            DC.AllowDBNull = True
            DC.DataType = System.Type.GetType("System.String")
            DC.ColumnName = "ProductSpecificationID"
            DT.Columns.Add(DC)


            DC = New DataColumn
            DC.AllowDBNull = True
            DC.DataType = System.Type.GetType("System.String")
            DC.ColumnName = "SpecName"
            DT.Columns.Add(DC)

            DC = New DataColumn
            DC.AllowDBNull = True
            DC.DataType = System.Type.GetType("System.String")
            DC.ColumnName = "SpecValue"
            DT.Columns.Add(DC)

            Dim MyFile As System.IO.FileInfo

            Dim I As Integer = 0

            MyFile = New FileInfo(ViewState("SPECPath"))
            If InStr(MyFile.Name, ".xls") Then


                sProduct.ProductID = ViewState("ProductID")
                If objProduct.GetProduct(sProduct) = True Then


                    Try

                        MyConnection = New System.Data.OleDb.OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0; " & _
                              "data source=" & MyFile.FullName & "; Extended Properties=Excel 8.0;")



                        MyCommand = New System.Data.OleDb.OleDbDataAdapter( _
                        "select * from [Sheet1$]", MyConnection)


                        DS = New System.Data.DataSet



                        MyCommand.Fill(DS)
                        MyConnection.Close()


                        DT = DS.Tables(0)

                        Dim CurrentSeriesID As Integer = -1

                        If DT.Rows.Count > 0 Then

                            DeleteAllProductSPECS(ViewState("ProductID"))


                            For Each DR In DT.Rows

                                Try
                                    If Len(DR(2)) > 0 Then

                                        'set values
                                        dsSpecs.UpdateParameters("ProductSpecificationID").DefaultValue = DR(0)
                                        dsSpecs.UpdateParameters("ProductSpecificationValue").DefaultValue = DR(2)

                                        'insert spec
                                        dsSpecs.Update()

                                        II += 1
                                        blnsuccess = True

                                    Else 'Bad Row in Spreadsheet

                                        I += 1

                                    End If

                                Catch ex As Exception
                                    'Bad Row in Spreadsheet
                                    I += 1
                                    blnsuccess = False


                                End Try
                            Next

                            If I > 0 Then
                                Me.lblReturnMsg.Text = II & " SPEC(s) were successfully added for this model.  " & I & " SPEC(s) were ignored!"
                                Me.pnlReturnMsg.Visible = True
                                Me.pnlError.Visible = False
                                UpdatePanelErrors.Update()

                            Else
                                Me.lblReturnMsg.Text = II & " SPEC(s) were successfully added for this model."
                                Me.pnlReturnMsg.Visible = True
                                Me.pnlError.Visible = False
                                UpdatePanelErrors.Update()
                            End If



                        Else
                            Me.lblError.Text = "There are no records in this table.  Make sure that the first tab in the spreadsheet is named 'Sheet1' and try again"
                            Me.pnlError.Visible = True
                            Me.pnlReturnMsg.Visible = False
                            UpdatePanelErrors.Update()

                        End If

                    Catch exx As Exception
                        Select Case Err.Number
                            Case 5

                                Me.lblError.Text = "No valid records were found.  Make sure that the first tab in the spreadsheet is named 'Sheet1' and try again"
                                Me.pnlError.Visible = True
                                Me.pnlReturnMsg.Visible = False
                                UpdatePanelErrors.Update()
                            Case Else
                                Me.lblError.Text = exx.Message
                                Me.pnlError.Visible = True
                                Me.pnlReturnMsg.Visible = False
                                UpdatePanelErrors.Update()
                        End Select

                    Finally
                        MyFile.Delete()

                        'Rebind data?
                        Me.gvParts.DataBind()
                    End Try

                Else 'Model not found
                End If
            End If
        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally
            objPart = Nothing
            objProduct = Nothing
            DS = Nothing
            DT = Nothing
            DR = Nothing
            DC = Nothing
        End Try
    End Sub

    Protected Sub btnUploadSPECS_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUploadSPECS.Click

        Dim bsuccess As Boolean

        Try
            ResetPanels()

            If txtSPECSFile.Value <> "" And txtSPECSFile.Value <> "test" Then
                bsuccess = Me.UploadSPECFile()

            End If

            If bsuccess Then
                LoadSPECSFromFile()
                'viewstate("BOMPath") 
            End If

            Me.gvProductSpecs.DataBind()
        Catch ex As Exception
            Me.lblError.Text = ex.Message
            Me.pnlError.Visible = True
            Me.pnlReturnMsg.Visible = False
            UpdatePanelErrors.Update()
        Finally

        End Try
    End Sub

    Protected Sub gvParts_RowDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gvParts.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                If IsNumeric(e.Row.DataItem("MSRP")) Then
                    CType(e.Row.FindControl("lblMSRP"), Label).Text = FormatCurrency(e.Row.DataItem("MSRP"), 2)
                End If

            End If

        Catch ex As Exception

        End Try


    End Sub

    Protected Sub fvAddPart_ItemInserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewInsertedEventArgs) Handles fvAddPart.ItemInserted
        Me.gvParts.DataBind()

        Me.lblReturnMsg.Text = "Part successfully added to model."
        Me.pnlReturnMsg.Visible = True
        Me.pnlError.Visible = False
        UpdatePanelErrors.Update()
    End Sub

    Protected Sub fvAddPart_ItemInserting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.FormViewInsertEventArgs) Handles fvAddPart.ItemInserting

        Dim SqlDataReader As SqlDataReader


        Try
            lblPartError.Text = ""

            'Make sure a part number was entered
            If CType(fvAddPart.FindControl("txtPartNumber"), TextBox).Text = "" Then
                Throw New Exception("Error adding Part -- Enter Part Number!")
            End If

            'If they entered a non-numeric warranty duration .. throw error
            If CType(fvAddPart.FindControl("txtWarrantyDuration"), TextBox).Text <> "" AndAlso Not IsNumeric(CType(fvAddPart.FindControl("txtWarrantyDuration"), TextBox).Text) Then
                Throw New Exception("Error adding Part -- Warranty Duration must be NUMERIC!")
            End If

            'If they entered a non-numeric quantity .. throw error
            If CType(fvAddPart.FindControl("txtQuantity"), TextBox).Text <> "" AndAlso Not IsNumeric(CType(fvAddPart.FindControl("txtQuantity"), TextBox).Text) Then
                Throw New Exception("Error adding Part -- Quantity must be NUMERIC!")
            End If


            dsPartSearch.DataSourceMode = SqlDataSourceMode.DataReader
            dsPartSearch.SelectParameters("PartNumber").DefaultValue = CType(fvAddPart.FindControl("txtPartNumber"), TextBox).Text
            SqlDataReader = CType(dsPartSearch.Select(DataSourceSelectArguments.Empty), SqlDataReader)


            If SqlDataReader.Read Then
                e.Values("PartID") = SqlDataReader("PartID")
            Else
                e.Cancel = True
                lblError.Text = "Part not found."
                pnlError.Visible = True
                UpdatePanelErrors.Update()

            End If


        Catch ex As Exception
            lblError.Text = ex.Message
            pnlError.Visible = True
            UpdatePanelErrors.Update()
        Finally
            SqlDataReader = Nothing
        End Try


    End Sub

    Protected Sub gvParts_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles gvParts.RowDeleted

        Me.lblReturnMsg.Text = "Part successfully deleted from model."
        Me.pnlReturnMsg.Visible = True
        Me.pnlError.Visible = False
        UpdatePanelErrors.Update()
    End Sub

    Protected Sub gvManuals_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles gvManuals.RowDeleted
        Try

        Catch ex As Exception

        End Try
    End Sub

    Protected Sub dsProductPromotion_Deleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsProductPromotion.Deleted

        Try
            DisplayMessage(MessageType.GeneralMsg, "Product was removed from selected location.")
            Me.UpdatePanelProductPromotion.Update()
        Catch ex As Exception
            DisplayMessage(MessageType.ErrorMsg, "There was a problem deleting the product from selected location.")
        End Try


    End Sub

    Protected Sub dsProductPromotion_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsProductPromotion.Inserted
        Try
            '   ViewState("Returnvalue") = e.Command.ExecuteScalar()
            Dim ReturnValue As Integer = 0

            ReturnValue = e.Command.Parameters("@ReturnValue").Value
            Select Case ReturnValue
                Case "0"
                    ViewState("Message") = "There was a problem adding product to featured location."
                    ViewState("Success") = "0"
                Case "-1"
                    ViewState("Message") = "This product already exists for the selected Location. "
                    ViewState("Success") = "0"
                Case Is > 1
                    ViewState("Message") = "Product was added successfully to the selected location. "
                    ViewState("Success") = "1"
            End Select


        Catch ex As Exception
            ViewState("Message") = "There was a problem adding product to promotion location."
            ViewState("Success") = "0"

        End Try
    End Sub

    Protected Sub dsProductPromotion_Inserting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) Handles dsProductPromotion.Inserting
        'TODO  - For testing only
        Try
            e.Command.Parameters("@CreatedBy").Value = CInt(Me.User.Identity.Name)
        Catch ex As Exception
            e.Command.Parameters("@CreatedBy").Value = 3746 'for testing on local machine
        End Try
    End Sub

    Protected Sub dsProductPromotion_Updated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsProductPromotion.Updated
        Try

            DisplayMessage(MessageType.GeneralMsg, "Product custom description was updated successfully. ")
            Me.UpdatePanelProductPromotion.Update()
        Catch ex As Exception
            DisplayMessage(MessageType.ErrorMsg, "There was a problem updating the custom description. ")
        Finally

        End Try
    End Sub



    Protected Sub dsProductPromotion_Updating(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) Handles dsProductPromotion.Updating
        'Set the ID of user updating description
        'TODO  - For testing only
        Try
            e.Command.Parameters("@UpdatedBy").Value = CInt(Me.User.Identity.Name)
        Catch ex As Exception
            e.Command.Parameters("@UpdatedBy").Value = 3746 'for testing on local machine
        End Try

    End Sub

    Protected Sub btnAddProductPromotion_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim bContinue As Boolean = True


        Try


            If Not CType(fvProductInfo.FindControl("IsSellableCheckBox"), CheckBox).Checked Then
                bContinue = False
                '
                Me.DisplayMessage(MessageType.ErrorMsg, "Product is not mark as a sellable item")
            End If

            If Not CType(fvProductInfo.FindControl("ActiveCheckBox"), CheckBox).Checked AndAlso bContinue = True Then
                bContinue = False
                Me.DisplayMessage(MessageType.ErrorMsg, "Product is not mark as an active item")
            End If


            If bContinue Then


                Me.dsProductPromotion.Insert()


                If ViewState("Success") = "1" Then
                    DisplayMessage(MessageType.GeneralMsg, "Product was added successfully under the selected location.")
                    Me.gvCurrentPromotions.DataBind()
                    Me.txtPromotionDescription.Text = ""

                    Me.UpdatePanelProductPromotion.Update()
                Else
                    DisplayMessage(MessageType.ErrorMsg, ViewState("Message"))
                End If

                Me.UpdatePanelProductPromotion.Update()

            End If

        Catch ex As Exception

        End Try
    End Sub

    Private Sub DisplayMessage(ByVal MessageType As MessageType, ByVal MessageString As String)

        Select Case MessageType
            Case Manager_ManageProduct.MessageType.ErrorMsg
                Me.lblError.Text = MessageString
                Me.lblReturnMsg.Text = ""
                Me.pnlReturnMsg.Visible = False
                Me.pnlError.Visible = True

            Case Manager_ManageProduct.MessageType.GeneralMsg

                Me.lblReturnMsg.Text = MessageString
                Me.pnlReturnMsg.Visible = True
                Me.lblError.Text = ""
                Me.pnlError.Visible = False

        End Select
        UpdatePanelErrors.Update()

    End Sub

    Private Enum MessageType
        GeneralMsg = 0
        ErrorMsg = 1
    End Enum

    Protected Sub UpdateSellable_OnClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim IsSellable As Boolean = CType(fvProductInfo.FindControl("IsSellableCheckBox"), CheckBox).Checked
        Dim ta As New dsDALTableAdapters.ProductTableAdapter()
        Dim RecordsAffected As Integer = ta.UpdateProductIsSellable(IsSellable, ViewState("ProductID"))

        If RecordsAffected = 1 Then
            'Successfully updated!
            Me.DisplayMessage(MessageType.GeneralMsg, _
                String.Format("Successfully set the ""Is Sellable"" attribute of product <strong>{0}</strong> to {1}.", _
                ViewState("ProductID"), IsSellable.ToString))
        Else
            'Problem updating...
            Me.DisplayMessage(MessageType.ErrorMsg, _
                String.Format("Error occurred setting the ""Is Sellable"" attribute of product <strong>{0}</strong> to {1}. Please try again!", _
                ViewState("ProductID"), IsSellable.ToString))
        End If
        ta.Dispose()
    End Sub

    Protected Sub UpdateActive_OnClick(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim IsActive As Boolean = CType(fvProductInfo.FindControl("ActiveCheckBox"), CheckBox).Checked
        Dim ta As New dsDALTableAdapters.ProductTableAdapter()
        Dim RecordsAffected As Integer = ta.UpdateProductIsActive(IsActive, ViewState("ProductID"))

        If RecordsAffected = 1 Then
            'Successfully updated!
            Me.DisplayMessage(MessageType.GeneralMsg, _
                String.Format("Successfully set the ""Active"" attribute of product <strong>{0}</strong> to {1}.", _
                ViewState("ProductID"), IsActive.ToString))
        Else
            'Problem updating...
            Me.DisplayMessage(MessageType.ErrorMsg, _
                String.Format("Error occurred setting the ""Active"" attribute of product <strong>{0}</strong> to {1}. Please try again!", _
                ViewState("ProductID"), IsActive.ToString))
        End If
        ta.Dispose()
    End Sub
End Class

