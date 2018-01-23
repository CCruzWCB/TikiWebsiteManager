Imports System.Xml
Imports System.Threading
Imports System.Data.SqlClient
Imports System.Data
Imports System.Configuration
Imports System.IO

Imports Refresh
Imports System.Collections.Generic

Partial Class Manager_ProductRefresh
    Inherits System.Web.UI.Page

    Shared bPublic As Boolean = True
    Shared rtRefreshType As RefreshType = RefreshType.PublicRefresh

    Private Enum RefreshType
        PublicRefresh = 1
        PrivateRefresh = 2
        RepPortalRefresh = 3

    End Enum


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then

            Me.Master.SetDisplayMessage("Select an option below to run.", MessageType.GeneralMessage)
            Me.Master.PageHeader = "Order Motion Refresh"
            Me.Master.PageTitle = "Order Motion Refresh"

            CheckStatus(Me, System.EventArgs.Empty)
        End If

    End Sub

    Public Function BizID() As String
        Dim strBizID As String

        strBizID = ConfigurationManager.AppSettings("OrderMotionHTTPBizID").ToString.Substring(0, 20) & "..."

        Return strBizID
    End Function



    Protected Sub btnRefreshRepPortal_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        'Add start entry in Refresh Log 
        LogWebStoreRefresh_START("Webstore Refresh - LL REP PORTAL")


        Me.gvRefreshLog.DataBind()
        upTimer.Update()


        'Turn Timer on
        Me.tRefreshTimer.Enabled = True

        'Start the refresh
        bPublic = False ' This is a refresh of the public web store
        rtRefreshType = RefreshType.RepPortalRefresh


        ThreadPool.QueueUserWorkItem(AddressOf RefreshWebStore)

        CheckStatus(Me, System.EventArgs.Empty)

    End Sub


    Protected Sub btnRefresh_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        'Add start entry in Refresh Log 
        If chkUpdateCategoriesPublic.Checked Then
            LogWebStoreRefresh_START("Webstore Refresh - PUBLIC - w/Categories")
        Else
            LogWebStoreRefresh_START("Webstore Refresh - PUBLIC")
        End If

        Me.gvRefreshLog.DataBind()
        upTimer.Update()


        'Turn Timer on
        Me.tRefreshTimer.Enabled = True

        'Start the refresh
        bPublic = True ' This is a refresh of the public web store
        rtRefreshType = RefreshType.PublicRefresh
        ThreadPool.QueueUserWorkItem(AddressOf RefreshWebStore)

        CheckStatus(Me, System.EventArgs.Empty)

    End Sub
    Protected Sub btnRefreshPrivate_Click(ByVal sender As Object, ByVal e As System.EventArgs)

        'Add start entry in Refresh Log 
        If chkUpdateCategoriesPrivate.Checked Then
            LogWebStoreRefresh_START("Webstore Refresh - PRIVATE - w/Categories")
        Else
            LogWebStoreRefresh_START("Webstore Refresh - PRIVATE")
        End If

        Me.gvRefreshLog.DataBind()
        upTimer.Update()


        'Turn Timer on
        Me.tRefreshTimer.Enabled = True

        'Start the refresh
        bPublic = False ' This is a refresh of the private web store
        ThreadPool.QueueUserWorkItem(AddressOf RefreshWebStore)

        CheckStatus(Me, System.EventArgs.Empty)

    End Sub

    Protected Sub btnRefreshShippingRates_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Add start entry in Refresh Log 
        LogWebStoreRefresh_START("Shipping Rate Refresh")
        Me.gvRefreshLog.DataBind()
        upTimer.Update()


        'Turn Timer on
        Me.tRefreshTimer.Enabled = True

        'Start the refresh
        bPublic = False ' This is a refresh of the private web store
        ThreadPool.QueueUserWorkItem(AddressOf RefreshShippingRates)

        CheckStatus(Me, System.EventArgs.Empty)

    End Sub

    Protected Sub CheckStatus(ByVal sender As Object, ByVal e As EventArgs) Handles tRefreshTimer.Tick
        Dim sqlDataReader As SqlDataReader = Nothing

        dsRefreshStatus.DataSourceMode = SqlDataSourceMode.DataReader
        sqlDataReader = CType(dsRefreshStatus.Select(DataSourceSelectArguments.Empty), SqlDataReader)

        If sqlDataReader.Read Then
            Me.btnRefreshPublic.Enabled = False
            Me.btnRefreshPrivate.Enabled = False
            btnRefreshShippingRates.Enabled = False
            upWebStoreRefreshButton.Update()
            Me.lblStatus.Text = sqlDataReader("RefreshStatus").ToString & " :: updated " & Now
            Me.lblStatus.ForeColor = Drawing.Color.Green
            hdnRefreshLogID.Value = sqlDataReader("RefreshLogID").ToString
            imgProcessing.Visible = True
        End If


        If InStr(lblStatus.Text, "Not Running") > 0 Then
            Me.lblStatus.ForeColor = System.Drawing.Color.FromName("#4097ca")


            Me.tRefreshTimer.Enabled = False
            imgProcessing.Visible = False
            Me.gvRefreshLog.DataBind()
            Me.btnRefreshPublic.Enabled = True
            Me.btnRefreshPrivate.Enabled = True
            btnRefreshShippingRates.Enabled = True
            upWebStoreRefreshButton.Update()
        End If
        upTimer.Update()

    End Sub
    Public Function LogWebStoreRefresh_START(Optional ByVal RefreshType As String = "Web Store") As Boolean

        Dim bSuccess As Boolean 'Return variable 

        'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

        Dim SQLCommand As New SqlCommand    'SQLCommand Object
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Try    'Only one "Try" statement 

            SQLConn.Open()    'Open Database


            'Set the Basic Command Information 
            SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
            SQLCommand.Connection = SQLConn       'Set the Connection


            'Set the Specific Command Information 
            SQLCommand.CommandText = "StartWebStoreRefresh"     'Stored Procedure Name

            'Stored Procedure Paramaters - 
            Try
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RefreshType", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, RefreshType))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CInt(Me.User.Identity.Name)))
            Catch ex As Exception
                'For testing in NON-PRODUCTION environment
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, 852))
            End Try




            SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


            'Close Connection
            SQLConn.Close()
            bSuccess = True

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
    Public Function LogWebStoreRefresh_END(ByVal Status As String, ByVal StatusNotes As String) As Boolean

        Dim bSuccess As Boolean 'Return variable 

        'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

        Dim SQLCommand As New SqlCommand    'SQLCommand Object
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Try    'Only one "Try" statement 

            SQLConn.Open()    'Open Database


            'Set the Basic Command Information 
            SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
            SQLCommand.Connection = SQLConn       'Set the Connection


            'Set the Specific Command Information 
            SQLCommand.CommandText = "EndWebStoreRefresh"     'Stored Procedure Name

            'Stored Procedure Paramaters - 
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RefreshLogID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, hdnRefreshLogID.Value))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Status", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Status))

            If StatusNotes <> "" Then
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@StatusNotes", System.Data.SqlDbType.VarChar, 250, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, StatusNotes))
            End If



            SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


            'Close Connection
            SQLConn.Close()
            bSuccess = True

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


    Private Sub RefreshWebStore(ByVal State As Object)
        Dim httpWebRequest As System.Net.HttpWebRequest
        Dim URLStreamWriter As System.IO.StreamWriter
        Dim httpWebResponse As System.Net.HttpWebResponse
        Dim URLStreamReader As System.IO.StreamReader
        Dim XMLString As String
        Dim WebStoreXML As XmlDocument
        Dim RequestXML As XmlDocument
        Dim StoresXml As XmlElement
        Dim StoreXml As XmlElement




        Dim blnContinue As Boolean = False

        'Try
        '    StoresXml = Me.GetWebStores()


        'Catch ex As Exception


        'Finally


        'End Try




        '  For Each StoreXml In StoresXml.ChildNodes

        blnContinue = False 'preset




        'This line for Public website
        '  If CType(StoreXml.Item("StoreCode").InnerText.Trim, Integer) OrElse CType(StoreXml.Item("StoreCode").InnerText.Trim, Integer) = 2 Then '<> "13" AndAlso StoreXml.Attributes("storeID").InnerText <> "1" Then

        Try
            RequestXML = New XmlDocument
            'RequestXML.Load("C:\OrderMotionXML\OrderMotion_GLCB_WebStoreInformationRequest.xml")
            ' RequestXML.LoadXml(String.Format("<?xml version='1.0' encoding='UTF-8' ?><WebStoreInformationRequest version='1.10'><UDIParameter><Parameter key='HTTPBizID'>{0}</Parameter><Parameter key='StoreID'>{1}</Parameter><Parameter key='OnlyItemsChangedSinceLastExported'>False</Parameter><Parameter key='OutputSubItemAttributes'>False</Parameter><Parameter key='ProductStatus'>All</Parameter></UDIParameter></WebStoreInformationRequest>", ConfigurationManager.AppSettings("OrderMotionHTTPBizID").ToString, StoreXml.Attributes("storeID").InnerText))

            Select Case rtRefreshType
                Case RefreshType.PublicRefresh
                    RequestXML.LoadXml(String.Format("<?xml version='1.0' encoding='UTF-8' ?><WebStoreInformationRequest version='1.10'><UDIParameter><Parameter key='HTTPBizID'>{0}</Parameter><Parameter key='StoreID'>{1}</Parameter><Parameter key='OnlyItemsChangedSinceLastExported'>False</Parameter><Parameter key='OutputSubItemAttributes'>False</Parameter><Parameter key='ProductStatus'>All</Parameter></UDIParameter></WebStoreInformationRequest>", ConfigurationManager.AppSettings("OrderMotionHTTPBizID").ToString, ConfigurationManager.AppSettings("OrderMotionStoreIDPUBLIC").ToString))
                Case RefreshType.PrivateRefresh
                    RequestXML.LoadXml(String.Format("<?xml version='1.0' encoding='UTF-8' ?><WebStoreInformationRequest version='1.10'><UDIParameter><Parameter key='HTTPBizID'>{0}</Parameter><Parameter key='StoreID'>{1}</Parameter><Parameter key='OnlyItemsChangedSinceLastExported'>False</Parameter><Parameter key='OutputSubItemAttributes'>False</Parameter><Parameter key='ProductStatus'>All</Parameter></UDIParameter></WebStoreInformationRequest>", ConfigurationManager.AppSettings("OrderMotionHTTPBizID").ToString, ConfigurationManager.AppSettings("OrderMotionStoreIDPRIVATE").ToString))
                Case RefreshType.RepPortalRefresh
                    RequestXML.LoadXml(String.Format("<?xml version='1.0' encoding='UTF-8' ?><WebStoreInformationRequest version='1.10'><UDIParameter><Parameter key='HTTPBizID'>{0}</Parameter><Parameter key='StoreID'>{1}</Parameter><Parameter key='OnlyItemsChangedSinceLastExported'>False</Parameter><Parameter key='OutputSubItemAttributes'>False</Parameter><Parameter key='ProductStatus'>All</Parameter></UDIParameter></WebStoreInformationRequest>", ConfigurationManager.AppSettings("OrderMotionHTTPBizID").ToString, ConfigurationManager.AppSettings("OrderMotionStoreIDREPPORTAL").ToString))
            End Select


            httpWebRequest = System.Net.HttpWebRequest.Create("https://api.omx.ordermotion.com/hdde/xml/udi.asp")
            httpWebRequest.Method = "POST"
            httpWebRequest.ContentType = "text/xml"
            httpWebRequest.Timeout = 400000
            URLStreamWriter = New IO.StreamWriter(httpWebRequest.GetRequestStream())


            'Write XML to OrderMotion
            Try
                URLStreamWriter.Write(RequestXML.OuterXml)
                URLStreamWriter.Close()


                'Read response from orderMotion
                Try


                    httpWebResponse = httpWebRequest.GetResponse()
                    URLStreamReader = New System.IO.StreamReader(httpWebResponse.GetResponseStream)
                    XMLString = URLStreamReader.ReadToEnd()
                    URLStreamReader.Close()

                    'Set the Item Information XML Document with the return xml
                    '    WebStoreXML = New Xml.XmlDocument
                    '    WebStoreXML.LoadXml(XMLString)
                    blnContinue = True



                Catch ex As Exception
                    '           _strError = String.Format("Could not read response from OrderMotion /Full Error -{0}", ex.Message)
                    blnContinue = False

                End Try

            Catch ex As Exception
                '        _strError = String.Format("Could not send information to OrderMotion /Full Error -{0}", ex.Message)
                blnContinue = False

            End Try


        Catch ex As Exception
            '_strError = String.Format("General Error - {0}", ex.Message)


        Finally
            httpWebRequest = Nothing
            URLStreamWriter = Nothing
            httpWebResponse = Nothing
            URLStreamReader = Nothing

        End Try

        'Corey, see me

        If blnContinue = True Then

            Select Case rtRefreshType

                Case RefreshType.PublicRefresh, RefreshType.PrivateRefresh
                    'MigrateParts_CB(XMLString)
                    If bPublic Then
                        If chkUpdateCategoriesPublic.Checked Then
                            Me.MigrateCategories_CB_Consumer(XMLString, bPublic)
                        End If
                    Else
                        If chkUpdateCategoriesPrivate.Checked Then
                            Me.MigrateCategories_CB_Consumer(XMLString, bPublic)
                        End If

                    End If

                    Me.MigrateItems_CB_Consumer(XMLString, bPublic)
                Case RefreshType.RepPortalRefresh

                    Me.MigrateItems_SalesReps(XMLString)


            End Select


        End If

        '  End If
        ' Next

    End Sub
    Private Sub RefreshShippingRates(ByVal State As Object)
        Dim httpWebRequest As System.Net.HttpWebRequest
        Dim URLStreamWriter As System.IO.StreamWriter
        Dim httpWebResponse As System.Net.HttpWebResponse
        Dim URLStreamReader As System.IO.StreamReader
        Dim XMLString As String
        Dim WebStoreXML As XmlDocument
        Dim RequestXML As XmlDocument
        Dim KeyCodesXml As XmlElement
        Dim KeyCodeXml As XmlElement

        Dim blnContinue As Boolean


        Try
            RequestXML = New XmlDocument
            'RequestXML.Load("C:\OrderMotionXML\OrderMotion_GLCB_WebStoreInformationRequest.xml")
            'RequestXML.LoadXml(String.Format("<?xml version='1.0' encoding='UTF-8' ?><KeycodeInformationRequest version='1.00'><UDIParameter><Parameter key='HTTPBizID'>{0}</Parameter><Parameter key='Keycode'>{1}</Parameter></UDIParameter></KeycodeInformationRequest>", ConfigurationManager.AppSettings("OrderMotionHTTPBizID").ToString, KeyCodeXml.Attributes("keycode").InnerText))
            RequestXML.LoadXml(String.Format("<KeycodeInformationRequest version='1.00'><UDIParameter><Parameter key='HTTPBizID'>{1}</Parameter><Parameter key='Keycode'>{0}</Parameter><Parameter key='DisplayKeycodeItems'>true</Parameter></UDIParameter></KeycodeInformationRequest>", "Consumer", ConfigurationManager.AppSettings("OrderMotionHTTPBizID").ToString))


            httpWebRequest = System.Net.HttpWebRequest.Create("https://api.omx.ordermotion.com/hdde/xml/udi.asp")
            httpWebRequest.Method = "POST"
            httpWebRequest.ContentType = "text/xml"
            httpWebRequest.Timeout = 400000
            URLStreamWriter = New IO.StreamWriter(httpWebRequest.GetRequestStream())


            'Write XML to OrderMotion
            Try
                URLStreamWriter.Write(RequestXML.OuterXml)
                URLStreamWriter.Close()


                'Read response from orderMotion
                Try


                    httpWebResponse = httpWebRequest.GetResponse()
                    URLStreamReader = New System.IO.StreamReader(httpWebResponse.GetResponseStream)
                    XMLString = URLStreamReader.ReadToEnd()
                    URLStreamReader.Close()

                    'Set the Item Information XML Document with the return xml
                    '    WebStoreXML = New Xml.XmlDocument
                    '    WebStoreXML.LoadXml(XMLString)
                    blnContinue = True



                Catch ex As Exception
                    '           _strError = String.Format("Could not read response from OrderMotion /Full Error -{0}", ex.Message)
                    blnContinue = False

                End Try

            Catch ex As Exception
                '        _strError = String.Format("Could not send information to OrderMotion /Full Error -{0}", ex.Message)
                blnContinue = False

            End Try


        Catch ex As Exception
            '_strError = String.Format("General Error - {0}", ex.Message)


        Finally
            httpWebRequest = Nothing
            URLStreamWriter = Nothing
            httpWebResponse = Nothing
            URLStreamReader = Nothing

        End Try


        If blnContinue = True Then
            Me.MigrateShippingRates(XMLString)
            'Me.MigratePrices(XMLString)
            'MigrateParts_CB(XMLString)
            ' Me.MigrateCategories_CB_Consumer(XMLString)
            'Me.MigrateItems_CB_Consumer(XMLString)
        End If
    End Sub

    Private Sub MigrateCategories_RepPortal(ByVal xmlDocument As XmlDocument)

        Dim xmlElement As System.Xml.XmlElement
        Dim sectionCategories As New List(Of SectionCategory)

        For Each xmlElement In xmlDocument.DocumentElement("WebStore").Item("Content").Item("SectionReference").ChildNodes
            Try
                Dim sectionCategory = PopulateSectionCategoryFromXmlElement(xmlDocument, xmlElement)
                If sectionCategory Is Nothing Then
                    Continue For
                End If

                sectionCategories.Add(sectionCategory)
            Catch ex As Exception

            Finally

            End Try
        Next

        AddOrUpdateSectionCategories(sectionCategories)

    End Sub

    Private Function PopulateSectionCategoryFromXmlElement(ByVal WebStoreXML As XmlDocument, ByVal xmlElement As XmlElement) As SectionCategory
        Dim sectionCategory As New SectionCategory
        Try
            With sectionCategory

                .Description = xmlElement.Item("Description").InnerText
                .SectionID = xmlElement.Attributes("sectionID").InnerText ' Webstore does NOT ues Categories
                If xmlElement.HasAttribute("catID") Then
                    .CategoryID = xmlElement.Attributes("catID").InnerText 'Can only used if webstore uses categories
                End If
                If xmlElement.HasAttribute("parentSectionID") Then
                    .ParentSectionID = xmlElement.Attributes("parentSectionID").InnerText 'Parent section ID, if this category is a sub-category
                End If
            End With
        Catch ex As Exception
            Return Nothing
        End Try
        Return sectionCategory
    End Function

    Private Function AddOrUpdateSectionCategories(ByVal sectionCategories As IEnumerable(Of SectionCategory)) As Boolean

        Dim bSuccess As Boolean 'Return variable 

        Using SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)
            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            Try
                SQLConn.Open()    'Open Database

                SQLCommand.CommandType = CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn

                SQLCommand.CommandText = "[LLRepPortal].[Refresh].[AddSectionCategory]"     'Stored Procedure Name

                For Each sectionCategory In sectionCategories
                    SQLCommand.Parameters.Clear()

                    With sectionCategory
                        SQLCommand.Parameters.Add(New SqlParameter("@SectionID", SqlDbType.Int, 8, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, .SectionID))

                        If .CategoryID Is Nothing OrElse String.IsNullOrEmpty(.CategoryID) Then
                            SQLCommand.Parameters.Add(New SqlParameter("@CategoryID", SqlDbType.Int, 8, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, DBNull.Value))
                        Else
                            SQLCommand.Parameters.Add(New SqlParameter("@CategoryID", SqlDbType.Int, 8, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, .CategoryID))
                        End If

                        If .Description Is Nothing OrElse String.IsNullOrEmpty(.Description) Then
                            SQLCommand.Parameters.Add(New SqlParameter("@Description", SqlDbType.VarChar, 255, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, DBNull.Value))
                        Else
                            SQLCommand.Parameters.Add(New SqlParameter("@Description", SqlDbType.VarChar, 255, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, .Description))
                        End If

                        If .ParentSectionID Is Nothing OrElse String.IsNullOrEmpty(.ParentSectionID) Then
                            SQLCommand.Parameters.Add(New SqlParameter("@ParentSectionID", SqlDbType.Int, 8, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, DBNull.Value))
                        Else
                            SQLCommand.Parameters.Add(New SqlParameter("@ParentSectionID", SqlDbType.Int, 8, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, .ParentSectionID))
                        End If

                        SQLCommand.Parameters.Add(New SqlParameter("@SectionCategoryID", SqlDbType.Int, 8, ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
                    End With

                    SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                    If Not SQLCommand.Parameters("@SectionCategoryID").Value > 0 Then
                        bSuccess = False
                    Else
                        sectionCategory.SectionCategoryID = SQLCommand.Parameters("@SectionCategoryID").Value

                        bSuccess = True
                    End If
                Next

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False

            End Try
        End Using


        'Set returning Boolean
        Return bSuccess
    End Function

    Private Function AddProductSectionCategories(ByVal productSectionCategories As List(Of sProductSectionCategory)) As Boolean
        Dim bSuccess As Boolean = True 'Return variable 

        Using SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)
            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            Try
                SQLConn.Open()    'Open Database

                SQLCommand.CommandType = CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn

                SQLCommand.CommandText = "[LLRepPortal].[Refresh].[AddProductSectionCategory]"     'Stored Procedure Name

                For Each productSectionCategory In productSectionCategories
                    SQLCommand.Parameters.Clear()

                    With productSectionCategory
                        SQLCommand.Parameters.Add(New SqlParameter("@ProductModelNumber", SqlDbType.VarChar, 50, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, productSectionCategory.ProductNumber))
                        SQLCommand.Parameters.Add(New SqlParameter("@SectionID", SqlDbType.Int, 8, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, .SectionID))
                    End With

                    SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure
                Next

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False

            End Try
        End Using


        'Set returning Boolean
        Return bSuccess
    End Function

    Private Function ClearProductSectionCategoryData() As Boolean
        Dim bSuccess As Boolean 'Return variable 

        Using SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)
            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            Try
                SQLConn.Open()    'Open Database

                SQLCommand.CommandType = CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn

                SQLCommand.CommandText = "[LLRepPortal].[Refresh].[ClearProductSectionCategory]"     'Stored Procedure Name

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False

            End Try
        End Using


        'Set returning Boolean
        Return bSuccess
    End Function

    Private Sub PopulateProductSectionCategoryFromXmlElement(ByRef ProductSectionCategories As List(Of sProductSectionCategory), ByRef xmlElement As XmlElement)
        Dim itemSections = xmlElement.Item("Section").GetElementsByTagName("ItemSection")

        For Each itemSection As XmlNode In itemSections
            Dim productSectionCategory As New sProductSectionCategory With {
                .ProductNumber = xmlElement.Attributes("itemCode").InnerText,
                .SectionID = itemSection.Attributes("sectionID").InnerText
            }
            ProductSectionCategories.Add(productSectionCategory)
        Next
    End Sub

    Private Sub MigrateCategories_CB_Consumer(ByVal WebStoreXmlString As String, ByVal bPublic As Boolean)
        'Add Categories

        Dim WebStoreXML As XmlDocument

        Dim xmlElement As System.Xml.XmlElement
        Dim objProductCategory As New ProductCategory
        Dim sProductCategory As sProductCategory
        Dim sParentProductCategory As sProductCategory
        Dim sOriginalProductCategory As sProductCategory







        Try

            WebStoreXML = New XmlDocument
            WebStoreXML.LoadXml(WebStoreXmlString)


            'First Delete all Category JUNC

            If objProductCategory.ResetAllEcomCategoryies(WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText) Then

                For Each xmlElement In WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference").ChildNodes

                    Try
                        With sProductCategory
                            .Active = True
                            .CreatedBy = 0
                            .DateCreated = Now()
                            .Name = xmlElement.Item("Description").InnerText
                            .Description = xmlElement.Item("Description").InnerText
                            .ExternalCategoryID = xmlElement.Attributes("sectionID").InnerText ' Webstore does NOT ues Categories
                            '.ExternalCategoryID = xmlElement.Attributes("catID").InnerText 'Can only used if webstore uses categories
                            .ExternalCategorySource = "ordermotion"
                            .PrimaryResourceImageID = 0
                            .BrandID = WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText

                        End With


                        'Check to see if category exist, if so update instead of add
                        sOriginalProductCategory = New sProductCategory
                        sOriginalProductCategory.ExternalCategoryID = sProductCategory.ExternalCategoryID
                        sOriginalProductCategory.BrandID = sProductCategory.BrandID
                        ' sOriginalProductCategory.ExternalCategorySource = "ordermotion"

                        If objProductCategory.GetProductCategory(sOriginalProductCategory) Then

                            sProductCategory.ProductCategoryID = sOriginalProductCategory.ProductCategoryID
                            objProductCategory.UpdateProductCategory(sProductCategory)
                        Else
                            objProductCategory.AddProductCategory(sProductCategory)

                        End If


                    Catch ex As Exception

                    Finally

                    End Try


                Next




                'Add JUNCs
                For Each xmlElement In WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference").ChildNodes
                    Try



                        'Get the category from BPS

                        sProductCategory = New sProductCategory
                        ' sProductCategory.ExternalCategoryID = xmlElement.Attributes("catID").InnerText 'WEBSTORE CATEGORY
                        sProductCategory.ExternalCategoryID = xmlElement.Attributes("sectionID").InnerText
                        sProductCategory.BrandID = WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText
                        sProductCategory.ExternalCategorySource = "ordermotion"

                        If objProductCategory.GetProductCategory(sProductCategory) Then


                            'Get the parent ID from OM (If needed)
                            Try
                                'GetCategoryID(xmlElement.Attributes("parentSectionID").InnerText, WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference"))
                                If CLng(xmlElement.Attributes("parentSectionID").InnerText) > 0 Then


                                    sParentProductCategory = New sProductCategory
                                    'If bPublic Then

                                    sParentProductCategory.ExternalCategoryID = xmlElement.Attributes("parentSectionID").InnerText 'GetCategoryID(xmlElement.Attributes("parentSectionID").InnerText, WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference"))
                                    'Else
                                    'sParentProductCategory.ExternalCategoryID = GetCategoryID(xmlElement.Attributes("parentSectionID").InnerText, WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference"))

                                    'End If


                                    'Reenable
                                    ' 
                                    sParentProductCategory.ExternalCategorySource = "ordermotion"
                                    sParentProductCategory.BrandID = WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText





                                    If objProductCategory.GetProductCategory(sParentProductCategory) Then
                                        sProductCategory.ParentCategoryID = sParentProductCategory.ProductCategoryID

                                    Else

                                        sProductCategory.ParentCategoryID = 0
                                    End If





                                Else
                                    sProductCategory.ParentCategoryID = 0


                                End If



                            Catch ex As Exception

                                sProductCategory.ParentCategoryID = 0

                            End Try


                            objProductCategory.AddProductCategoryBrandJunc(sProductCategory)



                        End If


                    Catch ex As Exception

                    Finally
                        '

                    End Try


                Next
            End If


        Catch ex As Exception


        Finally

            objProductCategory = Nothing
        End Try
    End Sub

    Private Sub MigrateItems_CB_Consumer(ByVal WebStoreXmlString As String, ByVal bPublic As Boolean)

        Dim WebStoreXML As XmlDocument
        Dim xmlItem As XmlElement
        Dim xmlDimension As XmlElement
        Dim xmlValue As XmlElement
        Dim xmlTEMP As XmlElement



        Dim xmlItemInfo As XmlNode
        Dim xmlItemInfos As XmlNodeList

        Dim xmlItemAttribute As XmlNode
        Dim xmlItemAttributes As XmlNodeList

        Dim xmlSubItem As XmlNode
        Dim xmlSubItems As XmlNodeList

        Dim xmlSubItemDimension As XmlNode
        Dim xmlSubItemDimensions As XmlNodeList

        Dim xmlAssociatedItems As XmlElement
        Dim xmlItemSections As XmlElement



        Dim blnItemHasAttributes As Boolean


        Dim objProductSeries As New ProductSeries
        Dim sProductSeries As sProductSeries
        Dim sOriginalProductSeries As sProductSeries
        Dim sProductSeriesAttribute As sProductSeriesAttribute

        Dim objProduct As New Product
        Dim sProduct As sProduct

        Dim blnHasSurcharges As Boolean = False
        Dim Surcharges As Double = 0.0
        Dim HighPrice As Double = 0.0
        Dim CurrentPrice As Double = 0.0

        Dim CurrentProductSeriesID As Long
        Dim CurrentProductID As Long

        Dim objProductCategory As New ProductCategory
        Dim sProductCategory As sProductCategory


        Dim strImage_ItemNumber As String = ""
        Dim BrandID As Integer = 0

        Try


            WebStoreXML = New XmlDocument
            WebStoreXML.LoadXml(WebStoreXmlString)

            BrandID = CInt(WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText)


            'Reset all "Issellable flags"

            Try
                objProductSeries.ResetAllEcomProducts(BrandID)
            Catch ex As Exception

            End Try


            For Each xmlItem In WebStoreXML.DocumentElement("WebStore").Item("Content").Item("ItemData").ChildNodes

                sProductSeries = New sProductSeries


                With sProductSeries
                    .Active = xmlItem.Attributes("active").InnerText
                    Try


                        .AdditionalShipping = xmlItem.Item("PriceData").Item("Price").Item("AdditionalSH").InnerText
                    Catch ex As Exception
                        .AdditionalShipping = 0.0

                    End Try

                    ' .Bonus = xmlItem.Item("PriceData").Item("Price").Item("Bonus").InnerText
                    .BrandID = BrandID  'WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText
                    '.CompanyID = 2
                    .CreatedBy = 0
                    .CreatedByName = "OM Integration service"
                    .WarrantyInfo = ""
                    .DateCreated = Now()
                    .IsSellable = .Active
                    xmlItemInfos = xmlItem.GetElementsByTagName("InfoText")

                    For Each xmlItemInfo In xmlItemInfos
                        If xmlItemInfo.Attributes("type").InnerText = "PlainText" Then .Description = xmlItemInfo.InnerText
                        If xmlItemInfo.Attributes("type").InnerText = "HTML" Then .WebDescription = xmlItemInfo.InnerText
                    Next

                    '  .Freight = ""

                    'no sub Items in CB
                    .HasSubProducts = False
                    'Check to see if item has subitems
                    'xmlSubItems = xmlItem.GetElementsByTagName("SubItem")
                    'If xmlSubItems.Count > 0 Then
                    '    .HasSubProducts = True
                    'End If


                    '  .IsBundle = xmlItem.Attributes("isBundle").InnerText


                    'Check to see if item is has attributes

                    Try
                        blnItemHasAttributes = False 'preset
                        xmlItemAttributes = xmlItem.Item("AttributeData").GetElementsByTagName("Attribute")
                        If xmlItemAttributes.Count > 0 Then
                            blnItemHasAttributes = True


                            'Check to see if this product has a product group and is "Instore only"

                            For Each xmlItemAttribute In xmlItemAttributes

                                Try


                                    Select Case LCase(xmlItemAttribute.Attributes("name").InnerText)

                                        Case "productgroup"
                                            If xmlItemAttribute.InnerText <> "" Then
                                                .ProductGroupName = xmlItemAttribute.InnerText
                                            End If


                                        Case "productgroupby"
                                            If xmlItemAttribute.InnerText <> "" Then
                                                .AttributeTypeName = xmlItemAttribute.InnerText

                                                Select Case LCase(.AttributeTypeName)

                                                    Case "size"

                                                        .AttributeTypeID = CInt(ProductSeriesAttributeType.Size)


                                                    Case "color"
                                                        .AttributeTypeID = CInt(ProductSeriesAttributeType.Color)

                                                End Select


                                            End If



                                        Case "instoreonly"

                                            If .IsSellable = True AndAlso xmlItemAttribute.InnerText <> "" AndAlso LCase(xmlItemAttribute.InnerText) = "true" Then
                                                .IsSellable = False
                                            End If


                                    End Select





                                Catch e As Exception


                                End Try



                            Next

                        Else

                            blnItemHasAttributes = False

                        End If

                    Catch ex As Exception
                        blnItemHasAttributes = False
                    End Try

                    'If blnItemHasAttributes = True Then
                    '    'Has attributes

                    '    ''Is it over sized?
                    '    'For Each xmlItemAttribute In xmlItemAttributes
                    '    '    If xmlItemAttribute.Attributes("name").InnerText = "Oversize Item" AndAlso LCase(xmlItemAttribute.InnerText) = "true" Then
                    '    '        .IsOversized = True
                    '    '        Exit For
                    '    '    End If
                    '    'Next
                    'End If


                    .IsSupported = True

                    ''Check to see if it is personalized
                    'Try

                    '    .IsPersonalized = xmlItem.Item("Personalization").Attributes("enabled").InnerText

                    '    If .IsPersonalized = True Then
                    '        .PersonalizationDescription = xmlItem.Item("Personalization").Attributes("label").InnerText
                    '    Else
                    '        .PersonalizationDescription = xmlItem.Item("Personalization").Attributes("label").InnerText
                    '        .PersonalizationDescription = ""

                    '    End If


                    'Catch ex As Exception
                    '    .PersonalizationDescription = ""
                    'End Try

                    ' .Multiplier = 0
                    .Name = xmlItem.Item("ProductName").InnerText


                    Try
                        .ExpectedDate = xmlItem.Item("NextExpectedDeliveryDate").InnerText

                    Catch ex As Exception
                        .ExpectedDate = Now.AddDays(28.0)

                    End Try


                    Try
                        .AvailableQuantity = xmlItem.Item("Available").InnerText

                    Catch ex As Exception
                        .AvailableQuantity = 0

                    End Try

                    'NextExpectedDeliveryDate

                    'New code to filter out "c" items
                    'strImage_ItemNumber = String.Format("{0}_{1}", .BrandID, xmlItem.Attributes("itemCode").InnerText)
                    strImage_ItemNumber = xmlItem.Attributes("itemCode").InnerText


                    'Get Primary Images(4)
                    If Me.GetItemPrimaryImage(strImage_ItemNumber, .BrandID.ToString, .PrimaryResourceImageIDFeature, .PrimaryResourceImageIDSmall, .PrimaryResourceImageID, .PrimaryResourceImageIDLarge) = False Then
                        .PrimaryResourceImageIDFeature = -3
                        .PrimaryResourceImageIDSmall = -2
                        .PrimaryResourceImageID = 0
                        .PrimaryResourceImageIDLarge = -1
                    End If


                    .ProductSeriesNumber = xmlItem.Attributes("itemCode").InnerText 'AS852
                    .Specifications = ""


                    'MSRP
                    'If it has sub products then we need to check to see if there is a price range
                    'if not then get single price
                    If .HasSubProducts = False Then
                        'single item
                        Try


                            If IsNumeric(xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText) Then
                                .MSRP = FormatCurrency(xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText, 2)
                            Else
                                .MSRP = "N/A"


                            End If
                        Catch ex As Exception
                            .MSRP = 0.0


                        End Try
                    Else
                        'No sub items in CB
                        ''sub items - need to check to see if there is a price range

                        ''First check to see if any dimension contains a surcharge


                        'For Each xmlDimension In xmlItem.Item("DimensionData").ChildNodes
                        '    For Each xmlValue In xmlDimension.ChildNodes
                        '        If xmlValue.Item("Surcharge").InnerText <> "" AndAlso xmlValue.Item("Surcharge").InnerText <> "0.00" Then
                        '            blnHasSurcharges = True
                        '            Exit For
                        '        End If

                        '    Next


                        'Next

                        'If blnHasSurcharges = False Then
                        '    If IsNumeric(xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText) Then
                        '        .MSRP = FormatCurrency(xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText, 2)
                        '    Else
                        '        .MSRP = "N/A"


                        '    End If

                        'Else


                        '    'Preset the High price 
                        '    HighPrice = 0.0

                        '    For Each xmlSubItem In xmlSubItems



                        '        'Preset the current price
                        '        CurrentPrice = xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText

                        '        'Get the SubItem dimensions
                        '        xmlTEMP = xmlSubItem 'TEMP CODE to convert node to element

                        '        xmlSubItemDimensions = xmlTEMP.GetElementsByTagName("ItemDimension")

                        '        'Now loop through the Sub Items dimensions to see if they have a surcharge

                        '        For Each xmlSubItemDimension In xmlSubItemDimensions

                        '            For Each xmlDimension In xmlItem.Item("DimensionData").ChildNodes

                        '                If xmlSubItemDimension.Attributes("name").InnerText = xmlDimension.Attributes("name").InnerText Then 'Attribute MATCHES!

                        '                    For Each xmlValue In xmlDimension.ChildNodes

                        '                        If xmlSubItemDimension.InnerText = xmlValue.Attributes("valueID").InnerText Then 'VALUE MATCHES!!
                        '                            If IsNumeric(xmlValue.Item("Surcharge").InnerText) AndAlso CType(xmlValue.Item("Surcharge").InnerText, Double) > 0.0 Then
                        '                                CurrentPrice += CType(xmlValue.Item("Surcharge").InnerText, Double)
                        '                            End If
                        '                            Exit For
                        '                        End If


                        '                    Next

                        '                    Exit For

                        '                End If


                        '            Next


                        '        Next 'Looping through SubItem Dimensions

                        '        'Now check to see if the current price is higher than the highest price
                        '        If CurrentPrice > HighPrice Then HighPrice = CurrentPrice


                        '    Next 'Looping through SubItems

                        '    'Finished looping throught the sub items, now check to see
                        '    'if there is a price range

                        '    If HighPrice > CDbl(xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText) Then
                        '        'PRICE RANGE
                        '        .MSRP = String.Format("{0} - {1}", FormatCurrency(xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText, 2), FormatCurrency(HighPrice, 2))
                        '    Else
                        '        'No price range
                        '        If IsNumeric(xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText) Then
                        '            .MSRP = FormatCurrency(xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText, 2)
                        '        Else
                        '            .MSRP = "N/A"


                        '        End If
                        '    End If

                        'End If


                    End If


                End With




                '*******************************************************************************************


                '
                'Now check to see if the product series already exist
                sOriginalProductSeries.ProductSeriesNumber = sProductSeries.ProductSeriesNumber
                '   sOriginalProductSeries.Specifications = ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString

                If objProductSeries.GetProductSeries(sOriginalProductSeries) Then

                    'EXISTING PRODUCT SERIES
                    sProductSeries.ProductSeriesID = sOriginalProductSeries.ProductSeriesID
                    sProductSeries.ProductID = sOriginalProductSeries.ProductID
                    'sProductSeries.PrimaryResourceImageID = sOriginalProductSeries.PrimaryResourceImageID
                    'sProductSeries.PrimaryResourceImageIDLarge = sOriginalProductSeries.PrimaryResourceImageIDLarge
                    'sProductSeries.PrimaryResourceImageIDSmall = sOriginalProductSeries.PrimaryResourceImageIDSmall


                    'Update Product Series

                    Try



                        For Each xmlItemSections In xmlItem.Item("Section").ChildNodes

                            'Get the Category
                            Try
                                'Temp code
                                If xmlItem.Item("Section").ChildNodes.Count > 1 Then
                                    Dim alertme As Boolean
                                    alertme = True
                                End If




                                objProductCategory = New ProductCategory
                                sProductCategory = New sProductCategory


                                'If bPublic Then
                                sProductCategory.ExternalCategoryID = xmlItemSections.Attributes("sectionID").InnerText 'Me.GetCategoryID(xmlItemSections.Attributes("sectionID").InnerText, WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference"))
                                'Else
                                '    sProductCategory.ExternalCategoryID = Me.GetCategoryID(xmlItemSections.Attributes("sectionID").InnerText, WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference"))
                                'End If



                                'REENABLE FOR OLD WEB STOREs
                                ' 
                                'FOR WEBSTORE CATEGORY

                                sProductCategory.BrandID = BrandID  'WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText

                                'sProductCategory.ExternalCategorySource = "ordermotion"

                                If objProductCategory.GetProductCategory(sProductCategory) Then
                                    sProductSeries.ProductCategoryID = sProductCategory.ProductCategoryID

                                    Exit For
                                    'Try
                                    '    objProductSeries.AddProductSeriesJUNC(sProductSeries)

                                    'Catch ex As Exception

                                    'End Try
                                End If


                            Catch ex As Exception

                            Finally
                                objProductCategory = Nothing

                            End Try



                        Next



                        If objProductSeries.UpdateProductSeries(sProductSeries) Then



                            'Not Needed for CB
                            ' ''Remove Categories
                            ''If objProductSeries.DeleteProductSeriesJUNC(, sProductSeries.ProductSeriesID) Then


                            ''    'Now add categories
                            ''    For Each xmlItemSections In xmlItem.Item("Section").ChildNodes



                            ''        'Get the Category
                            ''        Try

                            ''            objProductCategory = New BPS_BLv2.BPS_BL.ProductCategory
                            ''            sProductCategory = New BPS_BLv2.BPS_BL.sProductCategory



                            ''            sProductCategory.ExternalCategoryID = Me.GetCategoryID(xmlItemSections.Attributes("sectionID").InnerText, WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference"))
                            ''            sProductCategory.ExternalCategorySource = "ordermotion"

                            ''            If objProductCategory.GetProductCategory(sProductCategory) Then
                            ''                sProductSeries.ProductCategoryID = sProductCategory.ProductCategoryID

                            ''                Try
                            ''                    objProductSeries.AddProductSeriesJUNC(sProductSeries)

                            ''                Catch ex As Exception

                            ''                End Try
                            ''            End If


                            ''        Catch ex As Exception

                            ''        Finally
                            ''            objProductCategory = Nothing

                            ''        End Try



                            ''    Next

                            ''End If















                            'Remove ALL Attributes
                            If objProductSeries.DeleteProductAttributes(sProductSeries.ProductID) Then

                                'Add Attributes (if needed)
                                If blnItemHasAttributes = True Then
                                    sProductSeriesAttribute.ProductSeriesID = sProductSeries.ProductID

                                    'Add Weight attribute first

                                    If xmlItem.Item("Weight").InnerText <> "" Then

                                        sProductSeriesAttribute.AttributeName = "Weight"
                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Weight
                                        sProductSeriesAttribute.AttributeValue = xmlItem.Item("Weight").InnerText
                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                    End If




                                    For Each xmlItemAttribute In xmlItemAttributes

                                        Try


                                            Select Case LCase(xmlItemAttribute.Attributes("name").InnerText)

                                                Case "bulletlist"
                                                    If xmlItemAttribute.InnerText <> "" Then
                                                        objProductSeries.AddProductBullets(sProductSeries.ProductID, xmlItemAttribute.InnerText)
                                                    End If

                                                Case "length"
                                                    If xmlItemAttribute.InnerText <> "" Then

                                                        sProductSeriesAttribute.AttributeName = "Length"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Length
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                    End If

                                                Case "width"
                                                    If xmlItemAttribute.InnerText <> "" Then

                                                        sProductSeriesAttribute.AttributeName = "Width"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Width
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                    End If

                                                Case "height"
                                                    If xmlItemAttribute.InnerText <> "" Then

                                                        sProductSeriesAttribute.AttributeName = "Height"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Height
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                    End If

                                                Case "qtyper"
                                                    If xmlItemAttribute.InnerText <> "" Then

                                                        sProductSeriesAttribute.AttributeName = "Quantity per"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.QtyPer
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                    End If

                                                Case "scent"
                                                    If xmlItemAttribute.InnerText <> "" Then

                                                        sProductSeriesAttribute.AttributeName = "Scent"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Scent
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                    End If

                                                Case "oversized"
                                                    If xmlItemAttribute.InnerText <> "" AndAlso LCase(xmlItemAttribute.InnerText) = "true" Then

                                                        sProductSeriesAttribute.AttributeName = "Oversize Item"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Oversized
                                                        sProductSeriesAttribute.AttributeValue = "true"
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                    End If



                                                Case "usagelocation"

                                                    If xmlItemAttribute.InnerText <> "" Then
                                                        sProductSeriesAttribute.AttributeName = "Usage Location"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.UsageLocation
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                    End If

                                                Case "size"

                                                    If xmlItemAttribute.InnerText <> "" Then
                                                        sProductSeriesAttribute.AttributeName = "Size"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Size
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                    End If


                                                Case "color"

                                                    If xmlItemAttribute.InnerText <> "" Then
                                                        sProductSeriesAttribute.AttributeName = "Color"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Color
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                    End If


                                                Case "instoreonly"

                                                    If xmlItemAttribute.InnerText <> "" Then
                                                        sProductSeriesAttribute.AttributeName = "In Store Only"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.InStoreOnly
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                    End If

                                                Case "newproduct"

                                                    If xmlItemAttribute.InnerText <> "" Then
                                                        sProductSeriesAttribute.AttributeName = "New Product"
                                                        sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.NewProduct
                                                        sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                        objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                    End If

                                                Case "retailer"

                                                    If xmlItemAttribute.InnerText <> "" Then
                                                        objProductSeries.AddProductRetailers(sProductSeries.ProductID, xmlItemAttribute.InnerText)
                                                    End If
                                            End Select


                                        Catch ex As Exception

                                        End Try

                                    Next

                                End If


                            End If






                            'Disable ALL Sub Items

                            ''NOT Needed for CB ****************
                            ''If objProductSeries.DisableAllProducts(sProductSeries.ProductSeriesID) Then

                            ''    'No need to update the products if the the ProductSeries is disabled

                            ''    If sProductSeries.Active = True Then

                            ''        'Now check to see if it has sub items

                            ''        If sProductSeries.HasSubProducts Then

                            ''            For Each xmlSubItem In xmlSubItems


                            ''                With sProduct
                            ''                    'Get the available quantity
                            ''                    Try
                            ''                        .AvailableQuantity = xmlSubItem.Item("Available").InnerText

                            ''                    Catch ex As Exception
                            ''                        .AvailableQuantity = 0

                            ''                    End Try


                            ''                    .BrandID = sProductSeries.BrandID
                            ''                    .CompanyID = sProductSeries.CompanyID
                            ''                    .CreatedBy = sProductSeries.CompanyID
                            ''                    .CreatedByName = sProductSeries.CreatedByName
                            ''                    .DateCreated = Now()
                            ''                    .Description = sProductSeries.Description

                            ''                    Try
                            ''                        .ExpectedDate = xmlSubItem.Item("NextExpectedDeliveryDate").InnerText

                            ''                    Catch ex As Exception
                            ''                        .ExpectedDate = Now.AddDays(28.0)

                            ''                    End Try

                            ''                    .Freight = ""
                            ''                    .IsActive = sProductSeries.Active
                            ''                    .IsSellable = True
                            ''                    .IsSupported = True





                            ''                    Try
                            ''                        'preset Surcharges
                            ''                        Surcharges = 0.0

                            ''                        'preset msrp
                            ''                        .MSRP = xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText

                            ''                        'Preset name
                            ''                        .Name = sProductSeries.Name

                            ''                        'Now we need to loop through and get surcharges and Dimensions

                            ''                        'Get the SubItem dimensions
                            ''                        xmlTEMP = xmlSubItem 'TEMP CODE to convert node to element

                            ''                        xmlSubItemDimensions = xmlTEMP.GetElementsByTagName("ItemDimension")

                            ''                        'Now loop through the Sub Items dimensions to see if they have a surcharge

                            ''                        For Each xmlSubItemDimension In xmlSubItemDimensions

                            ''                            For Each xmlDimension In xmlItem.Item("DimensionData").ChildNodes

                            ''                                If xmlSubItemDimension.Attributes("name").InnerText = xmlDimension.Attributes("name").InnerText Then 'Attribute MATCHES!

                            ''                                    For Each xmlValue In xmlDimension.ChildNodes

                            ''                                        If xmlSubItemDimension.InnerText = xmlValue.Attributes("valueID").InnerText Then 'VALUE MATCHES!!
                            ''                                            .Name &= String.Format(" - {0}", xmlValue.Item("Description").InnerText)

                            ''                                            If IsNumeric(xmlValue.Item("Surcharge").InnerText) AndAlso CType(xmlValue.Item("Surcharge").InnerText, Double) > 0.0 Then
                            ''                                                Surcharges += CType(xmlValue.Item("Surcharge").InnerText, Double)
                            ''                                            End If
                            ''                                            Exit For
                            ''                                        End If


                            ''                                    Next

                            ''                                    Exit For

                            ''                                End If


                            ''                            Next


                            ''                        Next 'Looping through SubItem Dimensions




                            ''                    Catch ex As Exception

                            ''                    End Try


                            ''                    .PrimaryResourceImageID = sProductSeries.PrimaryResourceImageID
                            ''                    .PrimaryResourceImageIDLarge = sProductSeries.PrimaryResourceImageIDLarge
                            ''                    .PrimaryResourceImageIDSmall = sProductSeries.PrimaryResourceImageIDSmall
                            ''                    .ProductNumber = xmlSubItem.Attributes("itemCode").InnerText
                            ''                    .ProductSeriesID = sProductSeries.ProductSeriesID
                            ''                    .ProductYear = Now.Year.ToString
                            ''                    .MSRP += Surcharges
                            ''                    .Surcharge = Surcharges

                            ''                    Try
                            ''                        .UPCCode = xmlSubItem.Item("UPCCode").InnerText

                            ''                    Catch ex As Exception
                            ''                        .UPCCode = ""


                            ''                    End Try

                            ''                    Try
                            ''                        .Weight = xmlItem.Item("Weight").InnerText

                            ''                    Catch ex As Exception

                            ''                    End Try

                            ''                    .Year = .ProductYear


                            ''                End With




                            ''                'Now check to see if Product exists


                            ''                If objProduct.GetProductID(CurrentProductID, sProduct.ProductNumber, sProductSeries.BrandID) Then
                            ''                    'Existing product, UPDATE
                            ''                    sProduct.ProductID = CurrentProductID

                            ''                    If objProduct.UpdateProduct(sProduct) Then

                            ''                    End If



                            ''                Else
                            ''                    'NEW product ADD
                            ''                    If objProduct.AddProduct(sProduct) Then

                            ''                    End If
                            ''                End If




                            ''            Next 'Looping through SubItems




                            ''        Else
                            ''            'No sub products, just create one "Product"

                            ''            With sProduct
                            ''                'Get the available quantity
                            ''                Try
                            ''                    .AvailableQuantity = xmlItem.Item("Available").InnerText

                            ''                Catch ex As Exception
                            ''                    .AvailableQuantity = 0

                            ''                End Try


                            ''                .BrandID = sProductSeries.BrandID
                            ''                .CompanyID = sProductSeries.CompanyID
                            ''                .CreatedBy = sProductSeries.CompanyID
                            ''                .CreatedByName = sProductSeries.CreatedByName
                            ''                .DateCreated = Now()
                            ''                .Description = sProductSeries.Description

                            ''                Try
                            ''                    .ExpectedDate = xmlItem.Item("NextExpectedDeliveryDate").InnerText

                            ''                Catch ex As Exception
                            ''                    .ExpectedDate = Now.AddDays(28.0)

                            ''                End Try

                            ''                .Freight = ""
                            ''                .IsActive = sProductSeries.Active
                            ''                .IsSellable = True
                            ''                .IsSupported = True
                            ''                Try
                            ''                    .MSRP = xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText

                            ''                Catch ex As Exception

                            ''                End Try

                            ''                .Name = sProductSeries.Name
                            ''                .PrimaryResourceImageID = sProductSeries.PrimaryResourceImageID
                            ''                .PrimaryResourceImageIDLarge = sProductSeries.PrimaryResourceImageIDLarge
                            ''                .PrimaryResourceImageIDSmall = sProductSeries.PrimaryResourceImageIDSmall
                            ''                .ProductNumber = sProductSeries.ProductSeriesNumber
                            ''                .ProductSeriesID = sProductSeries.ProductSeriesID
                            ''                .ProductYear = Now.Year.ToString
                            ''                .Surcharge = 0.0

                            ''                Try
                            ''                    .UPCCode = xmlItem.Item("UPCCode").InnerText

                            ''                Catch ex As Exception
                            ''                    .UPCCode = ""


                            ''                End Try

                            ''                Try
                            ''                    .Weight = xmlItem.Item("Weight").InnerText

                            ''                Catch ex As Exception

                            ''                End Try

                            ''                .Year = .ProductYear


                            ''            End With


                            ''            'Now check to see if Product exists
                            ''            If objProduct.GetProductID(CurrentProductID, sProduct.ProductNumber, sProductSeries.BrandID) Then
                            ''                'Existing product, UPDATE
                            ''                sProduct.ProductID = CurrentProductID

                            ''                If objProduct.UpdateProduct(sProduct) Then

                            ''                End If



                            ''            Else
                            ''                'NEW product ADD
                            ''                If objProduct.AddProduct(sProduct) Then

                            ''                End If
                            ''            End If

                            ''        End If

                            ''    End If ' Checked to see if ProductSeries was Enabled

                            ''End If ' Disabled all products by series




                        End If

                    Catch ex As Exception

                    End Try






















                Else

                    'New Product SERIES

                    'Step 1 - Add product Series
                    'Step 2 - Determine if it has sub-items
                    'Step 3 - Add product (and attributes)
                    'Step 4 - Add additional sub-items (Products) if needed



                    'Add Product Series

                    Try

                        'Add Categories

                        For Each xmlItemSections In xmlItem.Item("Section").ChildNodes



                            'Get the Category
                            Try

                                'Temp code
                                If xmlItem.Item("Section").ChildNodes.Count > 1 Then
                                    Dim alertme2 As Boolean
                                    alertme2 = True
                                End If




                                objProductCategory = New ProductCategory
                                sProductCategory = New sProductCategory

                                'REENABLE for old

                                'If bPublic Then
                                sProductCategory.ExternalCategoryID = xmlItemSections.Attributes("sectionID").InnerText 'Me.GetCategoryID(xmlItemSections.Attributes("sectionID").InnerText, WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference"))
                                'Else
                                '    sProductCategory.ExternalCategoryID = Me.GetCategoryID(xmlItemSections.Attributes("sectionID").InnerText, WebStoreXML.DocumentElement("WebStore").Item("Content").Item("SectionReference"))
                                'End If




                                sProductCategory.BrandID = WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText

                                If objProductCategory.GetProductCategory(sProductCategory) Then
                                    sProductSeries.ProductCategoryID = sProductCategory.ProductCategoryID

                                    Exit For
                                    'Try
                                    '    objProductSeries.AddProductSeriesJUNC(sProductSeries)

                                    'Catch ex As Exception

                                    'End Try
                                End If


                            Catch ex As Exception

                            Finally
                                objProductCategory = Nothing

                            End Try



                        Next





                        If objProductSeries.AddProductSeries(sProductSeries) Then




                            'Add Attributes (if needed)
                            If blnItemHasAttributes = True Then
                                sProductSeriesAttribute.ProductSeriesID = sProductSeries.ProductID

                                For Each xmlItemAttribute In xmlItemAttributes

                                    Try


                                        Select Case LCase(xmlItemAttribute.Attributes("name").InnerText)

                                            Case "bullet"
                                                If xmlItemAttribute.InnerText <> "" Then

                                                    sProductSeriesAttribute.AttributeName = "Bullet"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Bullet
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                End If

                                            Case "length"
                                                If xmlItemAttribute.InnerText <> "" Then

                                                    sProductSeriesAttribute.AttributeName = "Length"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Length
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                End If

                                            Case "width"
                                                If xmlItemAttribute.InnerText <> "" Then

                                                    sProductSeriesAttribute.AttributeName = "Width"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Width
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                End If

                                            Case "height"
                                                If xmlItemAttribute.InnerText <> "" Then

                                                    sProductSeriesAttribute.AttributeName = "Height"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Height
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                End If

                                            Case "qtyper"
                                                If xmlItemAttribute.InnerText <> "" Then

                                                    sProductSeriesAttribute.AttributeName = "Quantity per"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.QtyPer
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                End If

                                            Case "scent"
                                                If xmlItemAttribute.InnerText <> "" Then

                                                    sProductSeriesAttribute.AttributeName = "Scent"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Scent
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                End If

                                            Case "oversized"
                                                If xmlItemAttribute.InnerText <> "" AndAlso LCase(xmlItemAttribute.InnerText) = "true" Then

                                                    sProductSeriesAttribute.AttributeName = "Oversize Item"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Oversized
                                                    sProductSeriesAttribute.AttributeValue = "true"
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)

                                                End If

                                            Case "usagelocation"

                                                If xmlItemAttribute.InnerText <> "" Then
                                                    sProductSeriesAttribute.AttributeName = "Usage Location"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.UsageLocation
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                End If

                                            Case "size"

                                                If xmlItemAttribute.InnerText <> "" Then
                                                    sProductSeriesAttribute.AttributeName = "Size"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Size
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                End If


                                            Case "color"

                                                If xmlItemAttribute.InnerText <> "" Then
                                                    sProductSeriesAttribute.AttributeName = "Color"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.Color
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                End If


                                            Case "instoreonly"

                                                If xmlItemAttribute.InnerText <> "" Then
                                                    sProductSeriesAttribute.AttributeName = "In Store Only"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.InStoreOnly
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                End If

                                            Case "newproduct"

                                                If xmlItemAttribute.InnerText <> "" Then
                                                    sProductSeriesAttribute.AttributeName = "New Product"
                                                    sProductSeriesAttribute.AttributeTypeID = ProductSeriesAttributeType.NewProduct
                                                    sProductSeriesAttribute.AttributeValue = xmlItemAttribute.InnerText
                                                    objProductSeries.AddProductSeriesAttribute(sProductSeriesAttribute)
                                                End If

                                            Case "retailer"

                                                If xmlItemAttribute.InnerText <> "" Then
                                                    objProductSeries.AddProductRetailers(sProductSeries.ProductID, xmlItemAttribute.InnerText)
                                                End If


                                        End Select





                                    Catch ex As Exception

                                    End Try

                                Next

                            End If

                            'Now check to see if it has sub items

                            'Not used for CB ******************
                            ''If sProductSeries.HasSubProducts Then


                            ''    For Each xmlSubItem In xmlSubItems


                            ''        With sProduct
                            ''            'Get the available quantity
                            ''            Try
                            ''                .AvailableQuantity = xmlSubItem.Item("Available").InnerText

                            ''            Catch ex As Exception
                            ''                .AvailableQuantity = 0

                            ''            End Try


                            ''            .BrandID = sProductSeries.BrandID
                            ''            .CompanyID = sProductSeries.CompanyID
                            ''            .CreatedBy = sProductSeries.CompanyID
                            ''            .CreatedByName = sProductSeries.CreatedByName
                            ''            .DateCreated = Now()
                            ''            .Description = sProductSeries.Description

                            ''            Try
                            ''                .ExpectedDate = xmlSubItem.Item("NextExpectedDeliveryDate").InnerText

                            ''            Catch ex As Exception
                            ''                .ExpectedDate = Now.AddDays(28.0)

                            ''            End Try

                            ''            .Freight = ""
                            ''            .IsActive = sProductSeries.Active
                            ''            .IsSellable = True
                            ''            .IsSupported = True





                            ''            Try
                            ''                'preset Surcharges
                            ''                Surcharges = 0.0

                            ''                'preset msrp
                            ''                .MSRP = xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText

                            ''                'Preset name
                            ''                .Name = sProductSeries.Name

                            ''                'Now we need to loop through and get surcharges and Dimensions

                            ''                'Get the SubItem dimensions
                            ''                xmlTEMP = xmlSubItem 'TEMP CODE to convert node to element

                            ''                xmlSubItemDimensions = xmlTEMP.GetElementsByTagName("ItemDimension")

                            ''                'Now loop through the Sub Items dimensions to see if they have a surcharge

                            ''                For Each xmlSubItemDimension In xmlSubItemDimensions

                            ''                    For Each xmlDimension In xmlItem.Item("DimensionData").ChildNodes

                            ''                        If xmlSubItemDimension.Attributes("name").InnerText = xmlDimension.Attributes("name").InnerText Then 'Attribute MATCHES!

                            ''                            For Each xmlValue In xmlDimension.ChildNodes

                            ''                                If xmlSubItemDimension.InnerText = xmlValue.Attributes("valueID").InnerText Then 'VALUE MATCHES!!
                            ''                                    .Name &= String.Format(" - {0}", xmlValue.Item("Description").InnerText)

                            ''                                    If IsNumeric(xmlValue.Item("Surcharge").InnerText) AndAlso CType(xmlValue.Item("Surcharge").InnerText, Double) > 0.0 Then
                            ''                                        Surcharges += CType(xmlValue.Item("Surcharge").InnerText, Double)
                            ''                                    End If
                            ''                                    Exit For
                            ''                                End If


                            ''                            Next

                            ''                            Exit For

                            ''                        End If


                            ''                    Next


                            ''                Next 'Looping through SubItem Dimensions




                            ''            Catch ex As Exception

                            ''            End Try


                            ''            .PrimaryResourceImageID = sProductSeries.PrimaryResourceImageID
                            ''            .PrimaryResourceImageIDLarge = sProductSeries.PrimaryResourceImageIDLarge
                            ''            .PrimaryResourceImageIDSmall = sProductSeries.PrimaryResourceImageIDSmall
                            ''            .ProductNumber = xmlSubItem.Attributes("itemCode").InnerText
                            ''            .ProductSeriesID = sProductSeries.ProductSeriesID
                            ''            .ProductYear = Now.Year.ToString
                            ''            .MSRP += Surcharges
                            ''            .Surcharge = Surcharges

                            ''            Try
                            ''                .UPCCode = xmlSubItem.Item("UPCCode").InnerText

                            ''            Catch ex As Exception
                            ''                .UPCCode = ""


                            ''            End Try

                            ''            Try
                            ''                .Weight = xmlItem.Item("Weight").InnerText

                            ''            Catch ex As Exception

                            ''            End Try

                            ''            .Year = .ProductYear


                            ''        End With

                            ''        If objProduct.AddProduct(sProduct) Then





                            ''        End If



                            ''    Next 'Looping through SubItems




                            ''Else
                            ''    'No sub products, just create one "Product"

                            ''    With sProduct
                            ''        'Get the available quantity
                            ''        Try
                            ''            .AvailableQuantity = xmlItem.Item("Available").InnerText

                            ''        Catch ex As Exception
                            ''            .AvailableQuantity = 0

                            ''        End Try


                            ''        .BrandID = sProductSeries.BrandID
                            ''        .CompanyID = sProductSeries.CompanyID
                            ''        .CreatedBy = sProductSeries.CompanyID
                            ''        .CreatedByName = sProductSeries.CreatedByName
                            ''        .DateCreated = Now()
                            ''        .Description = sProductSeries.Description

                            ''        Try
                            ''            .ExpectedDate = xmlItem.Item("NextExpectedDeliveryDate").InnerText

                            ''        Catch ex As Exception
                            ''            .ExpectedDate = Now.AddDays(28.0)

                            ''        End Try

                            ''        .Freight = ""
                            ''        .IsActive = sProductSeries.Active
                            ''        .IsSellable = True
                            ''        .IsSupported = True
                            ''        Try
                            ''            .MSRP = xmlItem.Item("PriceData").Item("Price").Item("Amount").InnerText

                            ''        Catch ex As Exception

                            ''        End Try

                            ''        .Name = sProductSeries.Name
                            ''        .PrimaryResourceImageID = sProductSeries.PrimaryResourceImageID
                            ''        .PrimaryResourceImageIDLarge = sProductSeries.PrimaryResourceImageIDLarge
                            ''        .PrimaryResourceImageIDSmall = sProductSeries.PrimaryResourceImageIDSmall
                            ''        .ProductNumber = sProductSeries.ProductSeriesNumber
                            ''        .ProductSeriesID = sProductSeries.ProductSeriesID
                            ''        .ProductYear = Now.Year.ToString
                            ''        .Surcharge = 0.0

                            ''        Try
                            ''            .UPCCode = xmlItem.Item("UPCCode").InnerText

                            ''        Catch ex As Exception
                            ''            .UPCCode = ""


                            ''        End Try

                            ''        Try
                            ''            .Weight = xmlItem.Item("Weight").InnerText

                            ''        Catch ex As Exception

                            ''        End Try

                            ''        .Year = .ProductYear


                            ''    End With

                            ''    If objProduct.AddProduct(sProduct) Then





                            ''    End If





                            ''End If




                        End If

                    Catch ex As Exception

                    End Try



                End If




            Next ' Next Item Element









            'ALL Items have been Added/Updated Now
            'go back through the items to add any crosssell and/or bundled items


            'First Add product groups
            If bPublic Then

                Try
                    objProductSeries.AddProductGroups(BrandID) ' (WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText)

                Catch ex As Exception

                End Try
            End If


            'Secind Delete ALL Associated Products
            Try
                objProductSeries.DeleteAllAssociatedProducts(WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText)

            Catch ex As Exception

            End Try

            For Each xmlItem In WebStoreXML.DocumentElement("WebStore").Item("Content").Item("ItemData").ChildNodes


                'First confirm that the Item was found in BPS

                sProductSeries = New sProductSeries
                sProductSeries.ProductSeriesNumber = xmlItem.Attributes("itemCode").InnerText

                If objProductSeries.GetProductSeries(sProductSeries) Then

                    CurrentProductID = sProductSeries.ProductID




                    'Any Cross Sells?
                    Try

                        If xmlItem.Item("CrossSell").ChildNodes.Count > 0 Then
                            For Each xmlAssociatedItems In xmlItem.Item("CrossSell").ChildNodes

                                sProductSeries = New sProductSeries
                                sProductSeries.ProductSeriesNumber = xmlAssociatedItems.Attributes("itemCode").InnerText

                                'Confirm that the cross Sell item exists in BPS
                                If objProductSeries.GetProductSeries(sProductSeries) Then

                                    'Add the CrossSell
                                    Try

                                        objProductSeries.AddProductAssociation(CurrentProductID, sProductSeries.ProductID, ProductSeriesAssociatedType.CrossSell)

                                    Catch ex As Exception

                                    End Try



                                End If






                            Next
                        End If

                    Catch ex As Exception

                    End Try




                    'Any Bundle Items?
                    Try

                        If xmlItem.Item("Bundle").ChildNodes.Count > 0 Then
                            For Each xmlAssociatedItems In xmlItem.Item("Bundle").ChildNodes

                                sProductSeries = New sProductSeries
                                sProductSeries.ProductSeriesNumber = xmlAssociatedItems.Attributes("itemCode").InnerText

                                'Confirm that the Bundle item exists in BPS
                                If objProductSeries.GetProductSeries(sProductSeries) Then

                                    'Add the bundle
                                    Try

                                        objProductSeries.AddProductAssociation(CurrentProductID, sProductSeries.ProductID, ProductSeriesAssociatedType.CrossSell)

                                    Catch ex As Exception

                                    End Try



                                End If






                            Next
                        End If

                    Catch ex As Exception

                    End Try




                End If ' Product Series Found?





            Next



            'Finally, disable empty categories

            Try
                'objProductSeries.DisableEmptyEcomCategoryies(WebStoreXML.DocumentElement("WebStore").Item("StoreCode").InnerText)
            Catch ex As Exception

            End Try


            'Update entry in Refresh Log (hdnRefreshLogID.value)
            LogWebStoreRefresh_END("Success", "")
        Catch ex As Exception
            Dim hj As String
            hj = ex.Message
            'Update entry in Refresh Log (hdnRefreshLogID.value)
            LogWebStoreRefresh_END("Failure", ex.Message)

        End Try










    End Sub

    Private Function GetCategoryID(ByVal SectionID As String, ByVal SectionsXML As XmlElement) As String

        Dim xmlSection As XmlElement
        Dim returnValue As String = ""

        For Each xmlSection In SectionsXML.ChildNodes
            If xmlSection.Attributes("sectionID").InnerText = SectionID Then
                returnValue = xmlSection.Attributes("catID").InnerText
                Exit For

            End If
        Next

        Return returnValue


    End Function

    Private Function GetItemPrimaryImage(ByVal ItemCode As String, ByVal BrandID As String, ByRef PrimaryIDFeature As Long, ByRef PrimaryIDSmall As Long, ByRef PrimaryIDRegular As Long, ByRef PrimaryIDLarge As Long) As Boolean


        Dim returnValue As Boolean = True



        'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

        Dim SQLCommand As New SqlCommand   'SQLCommand Object
        Dim SQLDataReader As SqlDataReader 'SQLDataReader
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        Try 'Only one "Try" statement 

            SQLConn.Open() 'Open Database

            'Set the Basic Command Information 
            SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
            SQLCommand.Connection = SQLConn                          'Set the Connection

            SQLCommand.CommandText = "LLFBPS..[r_SearchForProductSeriesPrimaryImages]"


            'Check to see if the BrandID was passed

            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesNumber", System.Data.SqlDbType.VarChar, 75, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ItemCode))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.VarChar, 2, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, BrandID))

            'Set the DataReader
            SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)

            'Preset
            PrimaryIDSmall = -2
            PrimaryIDFeature = -3
            PrimaryIDRegular = 0
            PrimaryIDLarge = -1

            Do While SQLDataReader.Read


                If InStr(SQLDataReader("Description"), "_regular") > 0 Then
                    PrimaryIDRegular = SQLDataReader("ImageID")
                End If


                If InStr(SQLDataReader("Description"), "_large") > 0 Then
                    PrimaryIDLarge = SQLDataReader("ImageID")
                End If


                If InStr(SQLDataReader("Description"), "_small") > 0 Then
                    PrimaryIDSmall = SQLDataReader("ImageID")
                End If


                If InStr(SQLDataReader("Description"), "_feature") > 0 Then
                    PrimaryIDFeature = SQLDataReader("ImageID")
                End If

            Loop




        Catch SQLex As SqlException
            returnValue = False
        Catch ex As Exception
            returnValue = False
        Finally
            If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
            SQLDataReader = Nothing
            SQLCommand = Nothing
            SQLConn = Nothing
        End Try


        Return returnValue


    End Function

    Private Sub MigrateShippingRates(ByVal KeycodeXmlString As String)

        Dim KeyCodeXML As XmlDocument
        Dim xmlItem As XmlElement


        Dim objPolicy As New ProductCategory


        Dim PriceListID As Integer = 0

        Dim blnContinue As Boolean = False


        KeyCodeXML = New XmlDocument
        KeyCodeXML.LoadXml(KeycodeXmlString)



        Try
            objPolicy.RefreshShippingRates(KeyCodeXML)

            'Update entry in Refresh Log (hdnRefreshLogID.value)
            LogWebStoreRefresh_END("Success", "")
        Catch ex As Exception

        End Try

    End Sub

    Public Function RefreshShippingRates(ShippingPolicyXML As XmlDocument) As Boolean

        Dim bSuccess As Boolean 'Return variable 
        Dim Tier As XmlElement
        Dim Charge As XmlElement
        Dim Policy As XmlElement



        'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

        Dim SQLCommand As New SqlCommand    'SQLCommand Object
        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Try    'Only one "Try" statement 

            SQLConn.Open()    'Open Database


            'Set the Basic Command Information 
            SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
            SQLCommand.Connection = SQLConn       'Set the Connection


            For Each Policy In ShippingPolicyXML.DocumentElement("Keycode").Item("ShippingPolicies").ChildNodes
                SQLCommand.Parameters.Clear()


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_DeleteShippingRatesByPolicyID]"     'Stored Procedure Name


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShippingPolicyID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CInt(Policy.Attributes("policyID").Value)))

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'bSuccess = True

                'Reset CMD for next SP
                SQLCommand.CommandText = "LLFBPS..[r_AddShippingCharge]"     'Stored Procedure Name


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShippingMethodID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, 0))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MaxValue", System.Data.SqlDbType.Decimal, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, 0))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Charge", System.Data.SqlDbType.Money, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, 0))




                'now loop through and set new rates
                For Each Tier In Policy.ChildNodes


                    If Tier.Name.ToUpper = "TIER" Then

                        SQLCommand.Parameters("@MaxValue").Value = Tier.Attributes("maxAmount").Value

                        For Each Charge In Tier.ChildNodes()
                            Try



                                SQLCommand.Parameters("@ShippingMethodID").Value = Charge.Attributes("SHCode").Value
                                SQLCommand.Parameters("@Charge").Value = Charge.InnerText
                                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                            Catch ex As Exception
                                LogWebStoreRefresh_END("Failure", ex.Message)
                                '_strErrorMessage = ex.ToString       'Set the Classes Error Message
                            End Try



                        Next
                    End If

                Next

            Next

            'Close Connection
            SQLConn.Close()

            'SQL ERROR CATCH
        Catch SQLErr As SqlException
            LogWebStoreRefresh_END("Failure", SQLErr.Message)
            '_strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
            '_intErrorNumber = SQLErr.Number       'Set the Classes Error Number
            bSuccess = False


            'MISC ERROR CATCH
        Catch Err As Exception
            LogWebStoreRefresh_END("Failure", Err.Message)
            '_strErrorMessage = Err.ToString       'Set the Classes Error Message
            '_intErrorNumber = 0       'Set the Classes Error Number
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

    Private Sub MigrateItems_SalesReps(ByVal WebStoreXmlString As String)

        Dim WebStoreXML As XmlDocument
        Dim xmlItem As XmlElement

        Dim xmlItemInfo As XmlNode
        Dim xmlItemInfos As XmlNodeList

        Dim xmlItemAttribute As XmlNode
        Dim xmlItemAttributes As XmlNodeList

        Dim objProductSeries As New ProductSeries
        Dim sProductSeries As sProductSeries

        Dim productSectionCategories As New List(Of sProductSectionCategory)

        Try


            WebStoreXML = New XmlDocument
            WebStoreXML.LoadXml(WebStoreXmlString)


            Try
                objProductSeries.ResetAllRepProducts()
                MigrateCategories_RepPortal(WebStoreXML)
            Catch ex As Exception
                LogWebStoreRefresh_END("Failure", ex.Message)
            End Try


            For Each xmlItem In WebStoreXML.DocumentElement("WebStore").Item("Content").Item("ItemData").ChildNodes

                sProductSeries = New sProductSeries


                With sProductSeries

                    .Name = xmlItem.Item("ProductName").InnerText
                    .ProductSeriesNumber = xmlItem.Attributes("itemCode").InnerText 'AS852
                    .Active = xmlItem.Attributes("active").InnerText
                    .DateCreated = Now()
                    .CaseQty = 1

                    xmlItemInfos = xmlItem.GetElementsByTagName("InfoText")

                    For Each xmlItemInfo In xmlItemInfos
                        If xmlItemInfo.Attributes("type").InnerText = "PlainText" Then .Description = xmlItemInfo.InnerText
                        If xmlItemInfo.Attributes("type").InnerText = "HTML" Then .WebDescription = xmlItemInfo.InnerText
                    Next

                    PopulateProductSectionCategoryFromXmlElement(productSectionCategories, xmlItem)

                    'Check to see if item is has attributes

                    Try

                        xmlItemAttributes = xmlItem.Item("AttributeData").GetElementsByTagName("Attribute")
                        If xmlItemAttributes.Count > 0 Then
                            For Each xmlItemAttribute In xmlItemAttributes
                                Try
                                    Select Case LCase(xmlItemAttribute.Attributes("name").InnerText)
                                        Case "oversized"
                                            If xmlItemAttribute.InnerText <> "" AndAlso LCase(xmlItemAttribute.InnerText) = "true" Then
                                                .IsOversized = True
                                            End If

                                        Case "caseqty"
                                            If xmlItemAttribute.InnerText <> "" AndAlso IsNumeric(xmlItemAttribute.InnerText) Then
                                                If Convert.ToInt64(xmlItemAttribute.InnerText) > 0 Then
                                                    .CaseQty = Convert.ToInt64(xmlItemAttribute.InnerText)
                                                End If
                                            End If
                                        Case "customerexclusive"
                                            If xmlItemAttribute.InnerText <> "" AndAlso LCase(xmlItemAttribute.InnerText) = "true" Then
                                                .IsHidden = True
                                            End If

                                        Case "caseqtyonly"
                                            If xmlItemAttribute.InnerText <> "" AndAlso LCase(xmlItemAttribute.InnerText) = "true" Then
                                                .IsCaseQTYOnly = True
                                            End If
                                    End Select
                                Catch ex As Exception

                                End Try
                            Next
                        End If

                    Catch ex As Exception

                    End Try
                End With

                Try

                    'Add Product
                    objProductSeries.AddProductToRepPortal(sProductSeries)

                Catch ex As Exception
                    LogWebStoreRefresh_END("Failure", ex.Message)
                End Try

            Next ' Next Item Element

            ClearProductSectionCategoryData()
            AddProductSectionCategories(productSectionCategories)

            LogWebStoreRefresh_END("Success", "")
        Catch ex As Exception

            'Update entry in Refresh Log (hdnRefreshLogID.value)
            LogWebStoreRefresh_END("Failure", ex.Message)



        End Try


    End Sub



End Class
