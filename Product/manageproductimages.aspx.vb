Imports Management
Imports BPS_BL.BPS
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO



Partial Class Product_manageproductimages
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

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            CleanUpOldImages() ' Attempt to cleanup obselete files 


            If Request.QueryString("ProductID") <> "" Then
                ViewState("ProductID") = Request.QueryString("ProductID")
            End If

            LoadImageProperites() ' Load 
            LoadProductImage()
            ToggleForm(FormMode.PrimaryImage)
        End If

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

    End Sub
    Public Sub CleanUpOldImages()
        'Clear this directory on page load. 
        Dim oTempDir As New DirectoryInfo(ConfigurationManager.AppSettings("TempResizeUploadDirectory").ToString)
        Dim oFile As FileInfo
        Try
            For Each oFile In oTempDir.GetFiles
                Try
                    oFile.Delete()
                Catch ex As Exception

                End Try
            Next
        Catch ex As Exception

        Finally
            oTempDir = Nothing
            oFile = Nothing
        End Try
        
        
    End Sub


  


    Private Sub ToggleForm(ByVal Mode As FormMode)
        Me.pPrimary.Visible = False
        Me.pAlternate.Visible = False

        Select Case Mode
            Case FormMode.PrimaryImage
                Me.lnkAlternate.CssClass = ""
                Me.lnkPrimary.CssClass = "itemselected"
                Me.pPrimary.Visible = True
                Me.Master.SetDisplayMessage("Welcome to Product Image Management", MessageType.GeneralMessage)

            Case FormMode.AlternateImages
                Me.lnkAlternate.CssClass = "itemselected"
                Me.lnkPrimary.CssClass = ""
                Me.pAlternate.Visible = True
                Me.gvAlternateImages.DataBind()

                Me.Master.SetDisplayMessage("Welcome to Alternate Product Image Management", MessageType.GeneralMessage)

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
            Return Utilities.GetImagePath(0)
        Finally

        End Try
    End Function

    Private Function UpdateProductSeriesImages(ByVal NewImageID As Integer, ByVal ImageType As MyImageSize) As Boolean
        Dim oProduct As New Product
        Dim sProduct As BPS_BL.BPS.sProduct = Nothing
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

            Me.Master.SetDisplayMessage("Error Updating Product Image Information: " & ex.Message.ToString, MessageType.ErrorMessage)

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
                Me.Master.SetDisplayMessage("Image " & NewFileName & " Updated Successfully ", MessageType.ConfirmationMessage)

            Else
                Throw New Exception("Error Adding/Updating Image: " & oImage.ErrorMessage)
            End If

        Catch ex As Exception
            AddImageFile = 0
            Me.Master.SetDisplayMessage("Error Adding/Updating Image: " & ex.Message.ToString, MessageType.ErrorMessage)
         
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
                Me.Master.SetDisplayMessage("Image " & NewFileName & " Association Updated Successfully ", MessageType.ConfirmationMessage)
                Me.gvAlternateImages.DataBind()

            Else
                Throw New Exception("Error Adding/Updating Image Association: " & oImage.ErrorMessage)
            End If

        Catch ex As Exception
            AddImageFiletoResource = False
            Me.Master.SetDisplayMessage("Error Adding/Updating Image Association: " & ex.Message.ToString, MessageType.ErrorMessage)
           

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

        ResourceTypeID = BPS_BL.BPS.ResourceType.Product


        
        'Get Upload Image Path to be used for adding or updating images
        ImagePath = Utilities.BuildImageUploadPath(ConfigurationManager.AppSettings("CompanyID"), sProduct.BrandID, ResourceTypeID)
        ' Response.Write(" Set Upload Image Path " & ImagePath)

        'Set Production Image File Location 
        ViewState("DefaultUploadPath") = ImagePath

        

        'Set temporary upload file location 
        ViewState("DefaultTempUploadPath") = ConfigurationManager.AppSettings("TempResizeUploadDirectory").ToString



        Try
            If oProduct.GetProduct(sProduct) Then
                'Set page title 
                'CP - Remove any "." from the product model before image file are created based on the model.
                ViewState("ProductModelNumber") = Replace(sProduct.ProductModelNumber, ".", "")

                Me.Master.SetDisplayMessage("Image Management for Product " & sProduct.Name, MessageType.GeneralMessage)
                Me.Master.PageHeader = sProduct.Name & " Images <BR>(" & sProduct.ProductModelNumber & ")"
                Me.Master.PageTitle = "Image Management"

                'Set the Product Series Number

                'Check for current primary image 
                If sProduct.PrimaryResourceImageID > 0 Then

                    Me.imgPrimary.ImageUrl = Utilities.GetImageURL(sProduct.ImagePath)
                    Me.imgRegular.ImageUrl = Utilities.GetImageURL(sProduct.ImagePath)
                    Me.Master.SetDisplayMessage("To change the primary image for this product, select the ""browse"" button below", Management.MessageType.instructionMessage)
                Else
                    Me.Master.SetDisplayMessage("To add a primary image, select the ""browse"" button below", Management.MessageType.instructionMessage)
                    Me.imgPrimary.ImageUrl = Utilities.GetImageURL(sProduct.ImagePath) 'Should be default no image 
                    Me.imgRegular.ImageUrl = Utilities.GetImageURL(sProduct.ImagePath)
                End If

                ViewState("ImageIDRegular") = sProduct.PrimaryResourceImageID


                'Check for current Small image 
                If sProduct.PrimaryResourceImageIDSmall > 0 Then
                    Me.Master.SetDisplayMessage("To change the primary image for this product, select the ""browse"" button below", Management.MessageType.instructionMessage)
                Else
                    Me.Master.SetDisplayMessage("To add a primary image, select the ""browse"" button below", Management.MessageType.instructionMessage)

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


                    Me.Master.SetDisplayMessage("To change the primary image for this product, select the ""browse"" button below", Management.MessageType.instructionMessage)
                Else
                    Me.Master.SetDisplayMessage("To add a primary image, select the ""browse"" button below", Management.MessageType.instructionMessage)
                    Me.imgLarge.ImageUrl = Utilities.GetImageURL(sProduct.ImagePathLarge)
                    'Me.imgLarge.ImageUrl = Utilities.GetImageURL(ConfigurationManager.AppSettings("NoImageURL"))
                End If

                'Store the current ImageIDs 
                ViewState("ImageIDLarge") = sProduct.PrimaryResourceImageIDLarge
                ViewState("ImageIDLargePath") = sProduct.ImagePathLarge



                'Check for current feature image 
                If sProduct.PrimaryResourceImageIDFeature > 0 Then
                    Me.Master.SetDisplayMessage("To change the feature image for this product, select the ""browse"" button below", Management.MessageType.instructionMessage)
                Else
                    Me.Master.SetDisplayMessage("To add a feature image, select the ""browse"" button below", Management.MessageType.instructionMessage)

                End If
                Me.imgFeature.ImageUrl = Utilities.GetImageURL(sProduct.ImagePathFeature)

                'Store the current ImageIDs 
                ViewState("ImageIDFeature") = sProduct.PrimaryResourceImageIDFeature
            Else
                Me.btnRegenerateImage.Visible = False
                Me.btnUpdateImage.Visible = False
                Me.chkAutoResize.Enabled = False

                Me.Master.SetDisplayMessage("This current Product is not available", Management.MessageType.ErrorMessage)

             
                Me.imgPrimary.ImageUrl = Utilities.GetImageURL(ConfigurationManager.AppSettings("NoImageURL"))

            End If

        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error Loading Image File" & ex.Message, MessageType.ErrorMessage)

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

        Dim bSuccess As Boolean = True

        'Upload largest size 
        If FileProductImageUpload.PostedFile.ContentLength > 0 And FileProductImageUpload.PostedFile.FileName <> "" Then

            If UploadFile(Me.FileProductImageUpload, MyImageSize.baseimage, True) >= 0 Then
                'store the original image path 
                ViewState("TempFile") = ViewState("DefaultFullUploadPath")


                'Create Large Image
                If Not Me.CreateResizedImage(ViewState("ProductID"), ViewState("DefaultFullUploadPath"), MyImageSize.large) Then
                    bSuccess = False
                End If

                'Create regular image
                If Not Me.CreateResizedImage(ViewState("ProductID"), ViewState("DefaultFullUploadPath"), MyImageSize.regular) Then
                    bSuccess = False
                End If

                'Create Small image
                If Not Me.CreateResizedImage(ViewState("ProductID"), ViewState("DefaultFullUploadPath"), MyImageSize.small) Then
                    bSuccess = False
                End If

                'Create featured image
                If Not Me.CreateResizedImage(ViewState("ProductID"), ViewState("DefaultFullUploadPath"), MyImageSize.feature) Then
                    bSuccess = False
                End If

            End If

            Try
                'TODO clean up base image
                'Note: this will probably failed due to locks held during previous resizing...
                'this director is cleared on page load. 
                CleanUpFile(ViewState("TempFile"), 0)



            Catch ex As Exception
                'This error is to be expected due to the windows/OS still having handle(lock) on file. 
                Me.Master.SetDisplayMessage("Error Cleaning up Temporary Files: " & ViewState("TempFile"), MessageType.ErrorMessage)
            End Try

            If bSuccess Then
                'Reload image 
                Me.LoadProductImage()
            End If


        Else
            Me.Master.SetDisplayMessage("Please select your Product Image using the ""Browse"" button.", MessageType.SyntaxMessage)
        End If


    End Sub



    Protected Sub btnAddAlternateImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddAlternateImage.Click
        Try
            Dim NewImageID As Integer = 0
            NewImageID = Me.UploadFile(Me.FileUploadAlternate, MyImageSize.Alternate)

            If NewImageID > 0 Then

                Dim FileName As String = Me.FileUploadAlternate.FileName
                ' Response.Write("Uploaded file  " & ViewState("DefaultFullUploadPath") & "<BR>")
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
            Me.Master.SetDisplayMessage("To generate small and regular product images, specify the image properties below for each image", MessageType.GeneralMessage)
        Else
            Me.Master.SetDisplayMessage("To change a specific product image size, select image file using the ""Browse"" button and click ""Update Image"" to save.", MessageType.GeneralMessage)
            Me.pImageSizeSettings.Visible = False
            Me.pProductImages.Visible = True
            Me.imgPrimary.Visible = False
        End If
    End Sub

    Protected Sub btnUpdateFeatureImage_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnUpdateFeatureImage.Click
        ResizeCurrentImage(Me.FileFeatureImageUpload, MyImageSize.feature)
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
            Me.Master.SetDisplayMessage("File was deleted successfully", MessageType.ConfirmationMessage)
        Else
            If Not IsNothing(e.Exception) Then
                Me.Master.SetDisplayMessage("Error deleting file: & " & e.Exception.Message.ToString, MessageType.ConfirmationMessage)
            Else
                Me.Master.SetDisplayMessage("There was a problem deleting file from web site", MessageType.ConfirmationMessage)
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
            Me.Master.SetDisplayMessage("Error updating Image: " & ex.Message.ToString, MessageType.ErrorMessage)
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
                Throw New Exception("Error Creating Small Featured Size")
            End If


            bSuccess = True

            If bSuccess Then
                Me.LoadProductImage()
            End If

        Catch ex As Exception




            Me.Master.SetDisplayMessage("There was a problem processing image: " & ex.Message.ToString, Management.MessageType.ErrorMessage)
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

        thumbnail = Nothing

    End Function

    Private Function ResizeAlternateImage(ByVal OriginalImageID As Integer, ByVal ImageUrl As String, ByVal ImageSize As MyImageSize) As Boolean
        Dim thumbNailImg As System.Drawing.Image
        Dim ResizedImg As System.Drawing.Image
        Dim fullSizeImg As System.Drawing.Image


        Try
            fullSizeImg = System.Drawing.Image.FromFile(ImageUrl)

            SetImageProperties(fullSizeImg, ImageSize)
            ' Response.Write("Image is about to be resized by object<BR>")
            ResizedImg = ImageResizer(fullSizeImg, imageWidth, imageHeight)
            ' Response.Write("Image has been resized..prepare to save to disk<BR>")

            Dim FullPath As String = ImagePath & NewFileName
            'Response.Write("Image save path = to : " & FullPath & "<BR>")
            'Response.Write("Image has been resize...now save to disk <BR>")
            ResizedImg.Save(FullPath, ImageFormat.Jpeg)
            'Response.Write("Image has been saved to : " & FullPath & "<BR>")


            Dim NewResizedImageID As Integer = AddImageFile(NewFileName, "Alternate Image View", ImagePath & NewFileName)

            If NewResizedImageID > 0 Then
                Me.AddImageFiletoResource(NewResizedImageID)
            Else
                Throw New Exception("Error saving resized alternate image")
            End If

            'Generate Thumbnail for image. 

            Dim NewThumbnailFileName As String = Utilities.GetFileName(NewFileName)
            NewThumbnailFileName = Utilities.parseProductNumberFromFileName(NewThumbnailFileName) & "_thumbnail.jpg"

            SetImageProperties(fullSizeImg, MyImageSize.AlternateThumb)
            '  Response.Write("Image thumbnail has been resized..prepare to save to disk<BR>")
            thumbNailImg = ImageResizer(fullSizeImg, imageWidth, imageHeight)
            FullPath = ImagePath & NewThumbnailFileName
            '  Response.Write("saved thumbnail to : " & FullPath & "<BR>")
            thumbNailImg.Save(FullPath, ImageFormat.Jpeg)
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
        SetImageProperties(fullSizeImg, ImageSize) 'Determine the size of the image by size

        Try
            If imageHeight > 0 And imageWidth > 0 Then

                thumbNailImg = ImageResizer(fullSizeImg, imageWidth, imageHeight)

                fullImagePath = ImagePath & NewFileName
                thumbNailImg.Save(fullImagePath, ImageFormat.Jpeg)


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
                fullSizeImg.Save(ImagePath & NewFileName, ImageFormat.Jpeg)
            End If

            bSuccess = True


        Catch ex As Exception
            bSuccess = False
            'Clean up / Dispose...
            thumbNailImg = Nothing
            fullSizeImg = Nothing
            
            Me.Master.SetDisplayMessage(ex.Message.ToString, MessageType.ErrorMessage)
        Finally

            thumbNailImg = Nothing
            fullSizeImg = Nothing


        End Try

        Return bSuccess

    End Function
    Sub SetImageProperties(ByVal ImageFile As System.Drawing.Image, ByVal ImageSize As MyImageSize)
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
                NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ConfigurationManager.AppSettings("BrandID") & "_" & ViewState("ProductModelNumber") & "_large.jpg"
            Case MyImageSize.feature
                If Not IsNothing(ImageFile) Then
                    If ImageFile.Height > imageHeight Then

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
                    End If
                Else
                    imageWidth = ConfigurationManager.AppSettings("MaxFeatureWidth")
                    imageHeight = ConfigurationManager.AppSettings("MaxFeatureHeight")
                End If

                ImagePath = ViewState("DefaultUploadPath")
                NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ConfigurationManager.AppSettings("BrandID") & "_" & ViewState("ProductModelNumber") & "_feature.jpg"

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
                NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ConfigurationManager.AppSettings("BrandID") & "_" & ViewState("ProductModelNumber") & "_regular.jpg"

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
                NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ConfigurationManager.AppSettings("BrandID") & "_" & ViewState("ProductModelNumber") & "_small.jpg"

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
                NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ConfigurationManager.AppSettings("BrandID") & "_" & ViewState("ProductModelNumber") & "_alternate_" & Left(System.Guid.NewGuid.ToString, 10) & ".jpg"

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
                NewFileName = ConfigurationManager.AppSettings("CompanyID") & "_" & ConfigurationManager.AppSettings("BrandID") & "_" & ViewState("ProductModelNumber") & "_alternate_thumb" & Left(System.Guid.NewGuid.ToString, 10) & ".jpg"

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

                'CP - changed based image name to generate randomly to help prevent file locks when uploading files back to back. 
                ImagePath = ViewState("DefaultUploadPath")
                NewFileName = Left(System.Guid.NewGuid.ToString, 10) & "_" & ViewState("ProductModelNumber") & "_baseimage.jpg"
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

            If (obj.PostedFile.ContentLength > 0 And obj.PostedFile.FileName.Length > 0) Then
                Dim ClientImage As System.Drawing.Image

                'Set image name, imageWidth, imageHeight, path & etc...
                SetImageProperties(ClientImage, ImageSize)
                fullPath = tempDirectory & NewFileName
                obj.PostedFile.SaveAs(fullPath)

            End If

            'Resize file to correct size based on imagesize 
            fullSizeImg = System.Drawing.Image.FromFile(fullPath)

            SetImageProperties(fullSizeImg, ImageSize)

            ResizedImg = ImageResizer(fullSizeImg, imageWidth, imageHeight)

            Dim ProductionFullPath As String = ImagePath & NewFileName

            'overwrite existing file
            ResizedImg.Save(ProductionFullPath, ImageFormat.Jpeg)

            'Clean up / Dispose...
            fullSizeImg.Dispose()

            ResizedImg.Dispose()

            fullSizeImg = Nothing
            ResizedImg = Nothing

            'delete temporary file 
            Try
                File.Delete(fullPath)
            Catch ex As Exception
                Me.Master.SetDisplayMessage("Error Removing temp file: " & fullPath & " - " & ex.Message.ToString, MessageType.ErrorMessage)

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
            Me.Master.SetDisplayMessage("Error resizing file: " & fullPath & " - " & ex.Message.ToString, MessageType.ErrorMessage)
          

        End Try



    End Sub


    Private Function UploadFile(ByRef obj As WebControls.FileUpload, ByVal ImageSize As MyImageSize, Optional ByVal UseTempDir As Boolean = False) As Integer
        Dim bSuccess As Boolean = False
        Dim NewPrimaryImageID As Integer = 0
        Dim BaseDirectoryPath As String = ""

        Try
            If (obj.PostedFile.ContentLength > 0 And obj.PostedFile.FileName.Length > 0) Then

                Dim ClientImage As System.Drawing.Image

                'Set image name, path & etc...
                SetImageProperties(ClientImage, ImageSize)

                
                If UseTempDir = True Then
                    ViewState("DefaultFullUploadPath") = ConfigurationManager.AppSettings("TempResizeUploadDirectory").ToString & NewFileName
                    BaseDirectoryPath = ConfigurationManager.AppSettings("TempResizeUploadDirectory").ToString
                    ViewState("TempFileName") = NewFileName
                Else
                    ViewState("DefaultFullUploadPath") = ViewState("DefaultUploadPath") & NewFileName
                    BaseDirectoryPath = ViewState("DefaultUploadPath")
                End If

                'Check to see if target Directory exists.
                If Not System.IO.Directory.Exists(BaseDirectoryPath) Then
                    System.IO.Directory.CreateDirectory(BaseDirectoryPath)
                End If

                obj.PostedFile.SaveAs(ViewState("DefaultFullUploadPath"))

                bSuccess = True

                'Create instance of image and check width to ensure dimensions are large ensure to be 
                'resized. 
                ClientImage = System.Drawing.Image.FromFile(ViewState("DefaultFullUploadPath"))

                Try
                    Dim oFile As New FileInfo(ViewState("DefaultFullUploadPath"))
                    'Make sure for furture update that file is not set to read only
                    If oFile.IsReadOnly Then
                        File.SetAttributes(ViewState("DefaultFullUploadPath"), FileAttributes.Normal)
                    End If

                    'Enforce Large when auto resizing config check is on and enabled. 
                    If ImageSize = MyImageSize.large And Me.chkAutoResize.Checked Then
                        If UCase(ConfigurationManager.AppSettings("Application_EnforceMaxWidthOnUpload")) = "ON" Then
                            If ClientImage.Width < ConfigurationManager.AppSettings("MaxLargeWidth") Then
                                Throw New Exception("Primary Product image must be atleast " & ConfigurationManager.AppSettings("MaxLargeWidth") & " Pixels in Width. Image " & NewFileName & " upload was aborted")
                            End If
                        End If
                    End If

                Catch ex As Exception

                Finally
                    ClientImage = Nothing
                End Try

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
                        '    Throw New Exception("There was a problem assigning New Product Image , Image upload has been aborted: " & NewFileName)
                    End If
                Else
                    '   Throw New Exception("There was a problem saving Primary Image to database, Image upload has been aborted: " & NewFileName)
                End If

            End If

        Catch ex As Exception
            Me.CleanUpFile(ViewState("DefaultFullUploadPath"), NewPrimaryImageID)
            NewPrimaryImageID = 0 'Return 0 is ImageID was not added. 
            Me.Master.SetDisplayMessage("Error Uploading File: " & ex.Message.ToString, MessageType.ErrorMessage)
        Finally

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

            'TODO for testing only...so not affect product data
                If FilePath <> "" Then
                File.Delete(FilePath)
            End If


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
        feature = 5
        Alternate = 3
        AlternateThumb = 6
        baseimage = 4
    End Enum
#End Region
    
 

    Protected Sub Page_Unload(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Unload
        
    End Sub
End Class
