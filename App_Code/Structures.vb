Imports System.Data
Imports System.Data.SqlClient
Imports system.Configuration



Namespace BPS_BL.BPS

    Public Structure sCompany
        Public CompanyID As Integer
        Public Name As String
        Public Description As String
        Public URL As String
        'Public LogoImage As String
        Public PrimaryResourceImageID As Integer
        Public PrimaryResourceImageName As String
        Public Phone As String
        Public ServiceLevelAgreement As String
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public IsActive As Boolean
        Public IsInternal As Boolean
        Public _LogInformation As String
        Public _UserInformation As String
    End Structure

    Public Structure sImage
        Public ImageID As Integer
        Public Description As String
        Public ImagePath As String
        Public Filename As String
        Public DateCreated As DateTime

        Public ImageResourceJunctionID As Integer
        Public ResourceTypeID As Integer
        Public ResourceType As String
        Public ResourceID As Integer
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
    End Structure
    Public Structure sProductDiagram
        Public ProductDiagramID As Integer
        Public ProductID As Integer
        Public Title As String
        Public Description As String
        Public FilePath As String
        Public FileName As String
        Public DateCreated As DateTime
        Public CreatedBy As Integer
        Public CreatedByName As String
    End Structure
    Public Structure sProduct
        Public ProductID As Long
        Public ProductModelNumber As String
        Public Name As String
        Public Description As String
        Public ProductYear As String
        Public Freight As String
        Public IsSupported As Boolean
        Public PrimaryResourceImageID As Long
        Public PrimaryResourceImageIDLarge As Long
        Public PrimaryResourceImageIDSmall As Long
        Public PrimaryResourceImageIDFeature As Long
        Public PrimaryResourceImageName As String
        Public PrimaryResourceImageNameLarge As String
        Public PrimaryResourceImageNameSmall As String
        Public PrimaryResourceImageNameFeature As String
        Public ImagePath As String
        Public ImagePathLarge As String
        Public ImagePathSmall As String
        Public ImagePathFeature As String
        Public Specifications As String
        Public MSRP As Double
        Public IsSellable As Boolean
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public ProductJunctionID As Integer
        Public ProductSeriesID As Integer
        Public ProductSeriesName As String
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public PriceListID As Integer
        Public ProductCategory1 As Integer
        Public ProductCategory1Name As String
        Public ProductCategory2 As Integer
        Public ProductCategory2Name As String
        Public DateCreated As Date
        Public IsActive As Boolean
        Public AvailableQuantity As Integer
        Public ExpectedDate As Date
        Public AdditionalShipping As Double
        Public _LogInformation As String
        Public _UserInformation As String
    End Structure
    Public Structure sProductSeries
        Public ProductSeriesID As Integer
        Public Name As String
        Public Description As String
        Public Freight As String
        Public IsSupported As Boolean
        Public Specifications As String
        Public PrimaryResourceImageID As Long
        Public PrimaryResourceImageIDLarge As Long
        Public PrimaryResourceImageIDSmall As Long
        Public PrimaryResourceImageIDFeature As Long
        Public PrimaryResourceImageName As String
        Public PrimaryResourceImageNameLarge As String
        Public PrimaryResourceImageNameSmall As String
        Public PrimaryResourceImageNameFeature As String
        Public ImagePath As String
        Public ImagePathLarge As String
        Public ImagePathSmall As String
        Public ImagePathFeature As String
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public DateCreated As Date
        Public Active As Boolean

        Public ProductSeriesJunctionID As Integer
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public ProductCategory1 As Integer
        Public ProductCategory1Name As String
        Public ProductCategory2 As Integer
        Public ProductCategory2Name As String

        Public _LogInformation As String
        Public _UserInformation As String
    End Structure

    Public Structure sCustomer
        Public CustomerName As String ' newly added cp 11/3/2008
        Public CustomerID As Long
        Public tempCustomerID As Long ' newly added cp 11/4/2008
        Public CustomerKey As String  ' newly added cp 11/3/2008
        Public FirstName As String
        Public MiddleName As String
        Public LastName As String
        Public FullName As String
        Public Gender As String
        Public Address1 As String
        Public Address2 As String
        Public City As String
        Public State As String
        Public Zip As String
        Public Country As String
        Public Email As String
        Public PrimaryPhone As String
        Public SecondaryPhone As String
        Public HomePhone As String
        Public WorkPhone As String
        Public MobilePhone As String
        Public FaxNumber As String
        Public AlternateCustomerNumber As String
        Public AlternateCustomerNumber2 As String
        Public AlternateCustomerNumber3 As String
        Public AlternateCustomerNumber4 As String
        Public UserField1 As String
        Public UserField2 As String
        Public UserField3 As String
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public CompanyID As Integer
        Public CompanyName As String
        Public ContactSheet As String
        Public BrandID As Integer
        Public BrandName As String
        Public AccountID As Integer
        Public AccountName As String
        Public CustomerAccountID As Long
        Public UserName As String
        Public Password As String
        Public DateCreated As Date
        Public Active As Boolean
        Public IsBusiness As Boolean
        Public TaxExemptNumber As String
        Public TaxStatusID As Integer
        Public RetailerID As String
        Public PriceListID As Integer  ' newly added cp 11/3/2008
        Public AccountStatusID As AccountStatus ' newly added cp 11/3/2008
        Public PurchaseStatusID As PurchaseStatus ' newly added cp 11/3/2008
        Public AllowMail As Boolean ' newly added cp 11/3/2008
        Public AllowEmail As Boolean  ' newly added cp 11/3/2008

        Public _LogInformation As String
        Public _UserInformation As String

        Public Function DisplayFullAddress() As String
            Dim fullAddress As String = ""
            fullAddress += Address1 & "<BR>"
            If Address2 <> "" Then '
                fullAddress &= "<BR>" & Address2 & "<BR>"
            End If
            fullAddress &= City & "," & State & " " & Zip
            fullAddress &= "<BR>" & Country
            Return fullAddress
        End Function

        Public Sub Load()

            Dim lngUserID As Long = 0 'Normalize Return Value
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand

            Try

                SQLCommand.CommandText = "LLFWebSitePrivate..[GetCustomerFullProfile]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))


                SQLConn.Open()
                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)

                If SQLDR.Read Then
                    CompanyName = SQLDR("CompanyName").ToString
                    CustomerName = SQLDR("Customer").ToString
                    FirstName = SQLDR("FirstName").ToString
                    LastName = SQLDR("LastName").ToString
                    Password = SQLDR("Password").ToString
                    Address1 = SQLDR("Address1").ToString
                    Address2 = SQLDR("Address2").ToString
                    IsBusiness = SQLDR("IsBusiness").ToString
                    City = SQLDR("City").ToString
                    State = SQLDR("State").ToString
                    Zip = SQLDR("Zip").ToString
                    Country = SQLDR("Country").ToString
                    AccountStatusID = SQLDR("AccountStatusID").ToString
                    PrimaryPhone = SQLDR("PrimaryPhone").ToString
                    SecondaryPhone = SQLDR("SecondaryPhone").ToString
                    CustomerID = SQLDR("CustomerID").ToString
                    Email = SQLDR("Email").ToString
                    AllowEmail = SQLDR("AllowEmail").ToString
                    AllowMail = SQLDR("AllowMail").ToString

                End If
            Catch SQLErr As SqlException
                'TODO create sys message for role not found 

            Catch ex As Exception

                'TODO create sys message for role not found 
            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLConn = Nothing
                SQLDR = Nothing
                SQLCommand = Nothing
            End Try

        End Sub

        Public Sub LoadfromCustomerID()

            Dim lngUserID As Long = 0 'Normalize Return Value
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand

            Try

                SQLCommand.CommandText = "LLFWebSitePrivate..[ValidateCustomerByID]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))
                'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerKey", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerKey))
                SQLConn.Open()
                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)

                If SQLDR.Read Then
                    CompanyName = SQLDR("CompanyName").ToString
                    CustomerID = SQLDR("CustomerID").ToString
                    Zip = SQLDR("Zip").ToString
                Else
                    Zip = ""
                    CustomerID = 0
                End If
            Catch SQLErr As SqlException
                'TODO create sys message for role not found 

            Catch ex As Exception

                'TODO create sys message for role not found 
            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLConn = Nothing
                SQLDR = Nothing
                SQLCommand = Nothing
            End Try

        End Sub

        Public Sub LoadfromCustomerTempCustomerID()

            Dim lngUserID As Long = 0 'Normalize Return Value
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand

            Try

                SQLCommand.CommandText = "LLFWebSitePrivate..[GetTempCustomerProfile]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@tempCustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, tempCustomerID))
                'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerKey", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerKey))
                SQLConn.Open()
                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)

                If SQLDR.Read Then

                    CustomerName = SQLDR("Customer").ToString
                    CustomerKey = SQLDR("CustomerKey").ToString
                    CompanyName = SQLDR("CompanyName").ToString
                    FirstName = SQLDR("FirstName").ToString
                    LastName = SQLDR("LastName").ToString
                    Address1 = SQLDR("Address1").ToString
                    Address2 = SQLDR("Address2").ToString
                    TaxExemptNumber = SQLDR("TaxExemptNumber").ToString

                    AlternateCustomerNumber = SQLDR("AlternateCustomerNumber").ToString
                    AlternateCustomerNumber2 = SQLDR("AlternateCustomerNumber2").ToString
                    AlternateCustomerNumber3 = SQLDR("AlternateCustomerNumber3").ToString
                    AlternateCustomerNumber4 = SQLDR("AlternateCustomerNumber4").ToString

                    City = SQLDR("City").ToString
                    State = SQLDR("State").ToString
                    Zip = SQLDR("Zip").ToString
                    Country = SQLDR("Country").ToString
                    AccountStatusID = SQLDR("AccountStatusID").ToString
                    PrimaryPhone = SQLDR("PrimaryPhone").ToString
                    SecondaryPhone = SQLDR("SecondaryPhone").ToString
                    CustomerID = SQLDR("CustomerID").ToString
                    Email = SQLDR("Email").ToString
                Else
                    tempCustomerID = 0
                End If
            Catch SQLErr As SqlException
                'TODO create sys message for role not found 

            Catch ex As Exception

                'TODO create sys message for role not found 
            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLConn = Nothing
                SQLDR = Nothing
                SQLCommand = Nothing
            End Try

        End Sub
    End Structure



    Public Structure sKnowledge
        Public KnowledgeID As Long
        Public KnowledgeCategoryID As Integer
        Public KnowledgeCategoryName As String
        Public KnowledgeTypeID As Integer
        Public KnowledgeTypeName As String
        Public Title As String
        Public Description As String
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public DateCreated As Date
        Public IsActive As Boolean
        Public IsFileOnly As Boolean
        Public CanPublishInternal As Boolean
        Public CanPublishExternal As Boolean
        Public IsQuickInfo As Boolean
        Public _LogInformation As String
        Public _UserInformation As String
    End Structure
    Public Structure sKnowledgeResource
        Public KnowledgeResourceID As Long
        Public ResourceTypeID As ResourceType
        Public ResourceID As Long
        Public ResourceName As String
        Public KnowledgeID As Long
        Public ResourceValue As String
        Public IsAlert As Boolean
        Public Active As Boolean
        Public CreatedBy As Long
        Public DateCreated As Date
    End Structure
    Public Structure sKnowledgeDocument
        Public KnowledgeDocumentID As Integer
        Public Name As String
        Public LocalSystemPath As String
        Public VirtualPath As String
        Public FullWebPath As String
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public DateCreated As Date


        'For KnowledgeDocumentJUNC
        Public KnowledgeID As Integer
        Public KnowledgeDocumentJUNCID As Integer

        Public _LogInformation As String
        Public _UserInformation As String

    End Structure

    Public Structure sPart
        Public PartID As Long
        Public Name As String
        Public Description As String
        Public PartNumber As String
        Public PrimaryResourceImageID As Integer
        Public PrimaryResourceImageName As String
        Public ImagePath As String
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public DateCreated As Date
        Public BaseWarrantyDuration As Integer
        Public PartCategoryID As Integer
        Public PartCategoryName As String
        Public CompanyID As Long
        Public BrandID As Long
        Public MSRP As Double
        Public IsSellable As Boolean


        'PartProductJUNC Fields
        Public PartProductID As Integer
        Public ProductID As Integer
        Public ProductModelNumber As String
        Public ProductName As String
        Public PartProductWarrantyDuration As Integer
        Public CallLetter As String

        'PartProductSeriesJUNC Fields
        Public PartProductSeriesID As Integer
        Public ProductSeriesID As Integer
        Public ProductSeriesName As String
        Public PartProductSeriesWarrantyDuration As Integer

        'Ecommerce Values
        Public ECommerce_Quantity As Integer
        Public ECommerce_Price As Integer
        Public ECommerce_PartIdentifier1 As String
        Public ECommerce_PartIdentifier2 As String
        Public ECommerce_PartIdentifier3 As String



        Public Active As Boolean

        Public _LogInformation As String
        Public _UserInformation As String
    End Structure

    Public Structure sSubscription
        Public SubscriptionID As Integer
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public Name As String
        Public Description As String
        Public DateCreated As Date
        Public IsActive As Boolean
        Public _LogInformation As String
        Public _UserInformation As String

    End Structure
    Public Structure sBrand
        Public BrandID As Integer
        Public CompanyID As Integer
        Public CompanyName As String
        Public Name As String
        Public Description As String
        Public URL As String
        Public PrimaryResourceImageID As Integer
        Public PrimaryResourceImageName As String
        'Public LogoImage As String
        Public Phone As String
        Public ServieLevelAgreement As String
        Public IsInternal As Boolean
        Public CreatedBy As Integer
        Public CreateByName As String
        Public DateCreated As Date
        Public IsActive As Boolean
        Public _LogInformation As String
        Public _UserInformation As String

    End Structure


    Public Structure sProductRegistration
        Public CustomerOwnershipID As Long
        Public ProductID As Long
        Public ProductModelNumber As String
        Public ProductName As String
        Public SerialNumber As String
        Public DateOfPurchase As Date
        Public PurchasePrice As String
        Public OwnershipConfirmed As Boolean
        Public CustomerID As Long
        Public ConfirmationSourceID As InteractionSource
        Public RegistrationSourceID As InteractionSource
        Public RetailerID As Long
        Public RetailerName As String
        Public DateCreated As Date
        Public DateConfirmed As Date
        Public IsActive As Boolean
        Public Reviewed As Boolean
        Public Customer As sCustomer
        Public Description As String

        'Fields for POP Requests
        Public FilePath As String
        Public FileName As String
        Public MimeType As String
        Public POPRequestedCaseID As Long
        Public POPRequestedDate As Date
        Public POPRequestedBy As Integer
        Public POPRequestedByName As String

        Public _LogInformation As String
        Public _UserInformation As String
    End Structure


    Public Structure sProductCategory
        Public ProductCategoryID As Integer
        Public Name As String
        Public Description As String
        Public ParentCategoryID As Integer
        Public ParentCategoryName As String
        Public PrimaryResourceImageID As Integer
        Public PrimaryResourceImageName As String
        Public PrimaryResourceImagePath As String
        Public PrimaryResourceImageFullWebPath As String
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public ChildrenCategories() As sProductCategory
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public DateCreated As Date
        Public Active As Boolean
        Public _LogInformation As String
        Public _UserInformation As String

    End Structure

    Public Structure sPaymentTerms
        Public CustomerID As Integer
        Public CustomerPaymentTermID As Integer
        Public PaymentTerms As String
        Public MaxInvoiceOrderAmount As Double
        Public MinInvoiceOrderAmount As Double
        Public MinInvoiceOrderAmountWeb As Double
        Public CurrentInvoicedAmount As Double
        Public NetTerms As Integer
        Public CreditLine As Double
        Public CreditApplicationStatusID As CreditApplicationStatus
        Public PurchaseStatusID As PurchaseStatus
        Public AvailableCreditLimit As Double


        Public Sub Load()

            Dim lngUserID As Long = 0 'Normalize Return Value

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLDR As SqlDataReader

            Dim SQLCommand As New SqlCommand

            Try

                '

                'SqlSelectCommand1

                '

                SQLCommand.CommandText = "LLFWebsitePrivate..[GetCustomerPaymentTerms]"

                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))



                SQLConn.Open()

                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)



                If SQLDR.Read Then

                    MaxInvoiceOrderAmount = SQLDR("MaxInvoiceOrderAmount").ToString
                    MinInvoiceOrderAmount = SQLDR("MinInvoiceOrderAmount").ToString

                    MinInvoiceOrderAmountWeb = SQLDR("MinInvoiceOrderAmountWeb").ToString

                    'this was replaced by available credit limit.
                    'CurrentInvoicedAmount = SQLDR("CurrentInvoicedAmount").ToString

                    NetTerms = SQLDR("NetTerms").ToString

                    CreditLine = SQLDR("CreditLine").ToString

                    CreditApplicationStatusID = SQLDR("CreditApplicationStatusID").ToString

                    PurchaseStatusID = SQLDR("PurchaseStatusID").ToString

                    AvailableCreditLimit = SQLDR("AvailableCreditLimit").ToString

                End If

            Catch SQLErr As SqlException

                'TODO create sys message for role not found 
                MaxInvoiceOrderAmount = 0
                MinInvoiceOrderAmount = 0
                MinInvoiceOrderAmountWeb = 0
                CurrentInvoicedAmount = 0
                NetTerms = 0
                CreditLine = 0
                CreditApplicationStatusID = CreditApplicationStatus.Unknown
                PurchaseStatusID = PurchaseStatus.NoInvoiceNoWeb

            Catch ex As Exception

                MaxInvoiceOrderAmount = 0
                MinInvoiceOrderAmount = 0
                MinInvoiceOrderAmountWeb = 0
                CurrentInvoicedAmount = 0
                NetTerms = 0
                CreditLine = 0
                CreditApplicationStatusID = CreditApplicationStatus.Unknown
                PurchaseStatusID = PurchaseStatus.NoInvoiceNoWeb

            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try

        End Sub

    End Structure


    Public Structure sPaymentOptions
        Public CustomerID As Integer
        Public AllowCreditCard As Boolean
        Public AllowInvoice As Boolean
        Public CurrentOrderSubTotal As Double
        Public CurrentOrderTotal As Double
        Public PaymentTerms As String
        Public DenyInvoiceReasonCode As Integer
        

        Public Sub Load()

            Dim lngUserID As Long = 0 'Normalize Return Value

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLDR As SqlDataReader

            Dim SQLCommand As New SqlCommand

            Try

                '

                'SqlSelectCommand1

                '

                SQLCommand.CommandText = "LLFWebsitePrivate..[GetCustomerPaymentOptions]"

                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SubTotal", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CurrentOrderSubTotal))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderTotal", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CurrentOrderTotal))


                SQLConn.Open()

                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)



                If SQLDR.Read Then

                    AllowInvoice = SQLDR("AllowInvoice")

                    AllowCreditCard = SQLDR("AllowCreditCard")

                    DenyInvoiceReasonCode = SQLDR("ReasonCode")

                End If

            Catch SQLErr As SqlException

                'TODO create sys message for role not found 
                AllowInvoice = 0

                AllowCreditCard = 0

                DenyInvoiceReasonCode = -1 'System Error 
            Catch ex As Exception

                'TODO create sys message for role not found 
                AllowInvoice = 0

                AllowCreditCard = 0

                DenyInvoiceReasonCode = -1 'System Error 
            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try

        End Sub

    End Structure

    Public Structure sAccount
        Public CustomerAccountID As Long
        Public AccountID As Long
        Public Name As String
        Public Description As String
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public DateCreated As Date
        Public IsActive As Boolean
        Public CustomerID As Long
        Public IsBusiness As Boolean
        Public PriceListID As Integer
        Public CustomerFirstname As String
        Public CustomerLastname As String
        Public CustomerZip As String
        Public AccountStatusID As AccountStatus
        Public Username As String
        Public Password As String
        Public PasswordExpirationDate As Date
        Public LastLoginDate As Date
        Public LoginDate As Date
        Public PurchaseStatusID As PurchaseStatus



        Public _LogInformation As String
        Public _UserInformation As String
        Public Function GetFullName() As String
            Return CustomerFirstname & " " & CustomerLastname
        End Function

        Public Sub SetLastLogin()
            
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand
            Try

                SQLCommand.CommandText = "LLFWebSitePrivate..[UpdateCustomerLastLoginDate]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))

                SQLConn.Open()

                SQLCommand.ExecuteNonQuery()


            Catch SQLErr As SqlException
                'TODO create sys message for role not found 
            Catch ex As Exception
                'TODO create sys message for role not found 
            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try
        End Sub

        Public Sub SetPassword()

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand
            Try

                SQLCommand.CommandText = "LLFBPS..[UpdateCustomerPassword]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Password", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Password))

                SQLConn.Open()

                SQLCommand.ExecuteNonQuery()


            Catch SQLErr As SqlException
                'TODO create sys message for role not found 
            Catch ex As Exception
                'TODO create sys message for role not found 
            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try
        End Sub
    End Structure


    Public Structure sCustomerOrder
        Public OrderID As Long
        Public CaseID As Integer
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public CustomerID As Long
        Public OrderNumber As String
        Public SubTotal As Double
        Public ShippingHandling As Double
        Public AdditionalSH As Double
        Public Tax As Double
        Public OrderTotal As Double
        Public ItemQuantity As Integer
        Public DateCreated As Date
        Public DateUpdated As Date
        Public DateClosed As Date
        Public CreatedBy As Long
        Public OrderSource As String
        Public Description As String
        Public TrackingNumber As String
        Public Active As Boolean
        Public ManualOrder As Boolean
        Public Customer As sCustomer
        Public ShoppingCart As String
        Public CreditCardInformation As String
        Public BillingAddress As String
        Public ShippingAddress As String
        Public PONumber As String
        Public StoreNumber As String


        Public _LogInformation As String
        Public _UserInformation As String


    End Structure

    Public Structure sCustomerOrderLineItem
        Public OrderLineItemID As Long
        Public OrderID As Long
        Public BpsItemType As String
        Public CaseID As Integer
        Public ItemNumber As String
        Public ItemDescription As String
        Public BpsID As Long
        Public ReasonCode As Integer
        Public ReasonCodeDescription As String
        Public ProblemModel As String
        Public Price As Double
        Public Tax As Double
        Public Quantity As Integer
        Public UnderWarranty As Boolean
        Public DateCreated As Date
        Public Status As Integer
        Public StatusDescription As String
        Public TrackingNumber As String
        Public TrackingSource As String
        Public LastUpdated As Date
        Public ShippingAddress As String
        Public DateShipped As Date





        Public _LogInformation As String
        Public _UserInformation As String


    End Structure


    Public Structure sRegisteredSerialNumber
        Public RegisteredSerialNumberID As Long
        Public CustomerOwnershipID As Long
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public CustomerID As Long
        Public CustomerName As String
        Public ProductID As Long
        Public ProductName As String
        Public ProductModelNumber As String
        Public SerialNumber As String
        Public SerialNumberConfirmed As Boolean
        Public DateCreated As Date
        Public _LogInformation As String
        Public _UserInformation As String


    End Structure

    Public Structure sCustomerNote
        Public CustomerNoteID As Long
        Public CustomerNoteTypeID As Integer
        Public CustomerNoteTypeName As String
        Public CustomerID As Long
        Public Customer As sCustomer
        Public Name As String
        Public Description As String
        Public CreatedBy As Integer
        Public CreatedByName As String
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public ProductID As Long
        Public ProductName As String
        Public ProductModelNumber As String
        Public DateCreated As Date
        Public IsActive As Boolean
        Public _LogInformation As String
        Public _UserInformation As String
    End Structure

    Public Structure sCustomerShipTo
        Dim CustomerID As Long
        Dim CustomerName As String
        Dim CustomerShipToID As Integer
        Dim ShipToNumber As Integer
        Dim AddressName As String
        Dim ContactName As String
        Dim Address1 As String
        Dim Address2 As String
        Dim City As String
        Dim State As String
        Dim Country As String
        Dim Zip As String
        Dim PhoneNumber As String
        Dim FaxNumber As String
        Dim CreatedBy As Integer
        Dim Active As Boolean
        Dim HasError As Boolean
        Dim ErrorMessage As String


        Public Sub LoadCustomerAddress()
            Dim lngUserID As Long = 0   'Normalize Return Value
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand

            Try

                '
                'SqlSelectCommand1
                '
                SQLCommand.CommandText = "LLFBPS..[GetCustomerAddress]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))

                SQLConn.Open()

                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)


                If SQLDR.Read Then
                    ShipToNumber = 0
                    CustomerName = SQLDR("FirstName").ToString & " " & SQLDR("LastName").ToString
                    Address1 = SQLDR("Address1").ToString
                    Address2 = SQLDR("Address2").ToString
                    City = SQLDR("City").ToString
                    State = SQLDR("State").ToString
                    Zip = SQLDR("Zip").ToString
                    Country = SQLDR("Country").ToString
                End If

            Catch SQLErr As SqlException
                'TODO create sys message for role not found 
            Catch ex As Exception
                'TODO create sys message for role not found 
            Finally
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing
                SQLDR = Nothing
                SQLCommand = Nothing
            End Try
        End Sub

        Public Sub LoadStoreAddressInfo()
            Dim lngUserID As Long = 0   'Normalize Return Value
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand

            Try

                '
                'SqlSelectCommand1
                '
                SQLCommand.CommandText = "LLFBPS..[GetCustomerShipToInfo]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerShipToID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerShipToID))

                SQLConn.Open()

                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)


                If SQLDR.Read Then

                    ShipToNumber = SQLDR("ShipToNumber").ToString
                    CustomerName = SQLDR("CustomerName").ToString
                    ContactName = SQLDR("ContactName").ToString
                    Address1 = SQLDR("Address1").ToString
                    Address2 = SQLDR("Address2").ToString
                    City = SQLDR("City").ToString
                    State = SQLDR("State").ToString
                    Country = SQLDR("Country").ToString
                    Zip = SQLDR("Zip").ToString
                    AddressName = SQLDR("Name").ToString
                    

                End If

            Catch SQLErr As SqlException
                'TODO create sys message for role not found 
                ErrorMessage = SQLErr.Message.ToString
                HasError = True
                ShipToNumber = -1
            Catch ex As Exception
                'TODO create sys message for role not found 
                ErrorMessage = ex.Message.ToString
                HasError = True
                ShipToNumber = -1
            Finally
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing
                SQLDR = Nothing
                SQLCommand = Nothing
            End Try

        End Sub

        Public Sub AddCustomerShiptoAddressInfo()
            Dim lngUserID As Long = 0   'Normalize Return Value
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand

            Try

                '
                'SqlSelectCommand1
                '
                SQLCommand.CommandText = "LLFWebSitePrivate..[AddCustomerShipTo]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShipToNumber", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ShipToNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ContactName", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ContactName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address1", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Address1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address2", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Address2))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@City", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, City))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@State", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, State))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Zip", System.Data.SqlDbType.VarChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Zip))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Country", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Country))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CreatedBy))
                '                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerShipToID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, DBNull.Value))

                SQLConn.Open()

                CustomerShipToID = SQLCommand.ExecuteScalar

                If CustomerShipToID > 0 Then

                ElseIf CustomerShipToID = -1 Then
                    CustomerShipToID = 0
                    ErrorMessage = "Ship-To Address already exists."
                    HasError = True
                ElseIf CustomerShipToID = 0 Then
                    CustomerShipToID = 0
                    ErrorMessage = "There was a problem adding the Ship-To information."
                    HasError = True
                End If

            Catch SQLErr As SqlException
                'TODO create sys message for role not found 
                CustomerShipToID = 0
                ErrorMessage = SQLErr.Message.ToString
                HasError = True
            Catch ex As Exception
                'TODO create sys message for role not found 
                CustomerShipToID = 0
                ErrorMessage = ex.Message.ToString
                HasError = True
            Finally
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing
                SQLDR = Nothing
                SQLCommand = Nothing
            End Try

        End Sub

        Public Sub UpdateCustomerShiptoAddressInfo()
            Dim lngUserID As Long = 0   'Normalize Return Value
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLDR As SqlDataReader
            Dim SQLCommand As New SqlCommand

            Try

                '
                'SqlSelectCommand1
                '
                SQLCommand.CommandText = "LLFWebSitePrivate..[UpdateCustomerShipTo]"
                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerShipToID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerShipToID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShipToNumber", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ShipToNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ContactName", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ContactName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address1", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Address1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address2", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Address2))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@City", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, City))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@State", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, State))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Zip", System.Data.SqlDbType.VarChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Zip))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Country", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Country))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@LastUpdatedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CreatedBy))

                SQLConn.Open()

                CustomerShipToID = SQLCommand.ExecuteScalar

                If CustomerShipToID > 0 Then

                ElseIf CustomerShipToID = -1 Then
                    CustomerShipToID = 0
                    ErrorMessage = "ShipTo Address already exists."
                    HasError = True
                ElseIf CustomerShipToID = 0 Then
                    CustomerShipToID = 0
                    ErrorMessage = "There was a problem adding the ShipTo information."
                    HasError = True
                End If

            Catch SQLErr As SqlException
                'TODO create sys message for role not found 
                CustomerShipToID = 0
                ErrorMessage = SQLErr.Message.ToString
                HasError = True
            Catch ex As Exception
                'TODO create sys message for role not found 
                CustomerShipToID = 0
                ErrorMessage = ex.Message.ToString
                HasError = True
            Finally
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing
                SQLDR = Nothing
                SQLCommand = Nothing
            End Try

        End Sub
    End Structure
    Public Structure sEmail

        '"sEmail" Structure Profile ******************************************************

        '

        ' Structure: sEmail

        ' Parent File: Structures.vb

        ' Parent Project: Business Logic

        ' Parent Solution: HLGIT CRM System

        ' Created By: Vincent Clover

        ' Created on: 07/02/04

        ' Description: Can Hold information of a Email

        '

        '

        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        Public MailToAddress As String

        Public MailFromAddress As String

        Public Subject As String

        Public Body As String

        Public SMTPServer As String

        Public Priority As MailPriority

        Public IsBodyHtml As Boolean

        Public ReplyTo As String

        Public Attachments As Collection

        Public TemplateID As EmailTemplate

        Public Title As String

        ' Public Origin website or cms





        'Used to contain plain text email content version
        Public TextBody As String

        'Used to determine whether marketing content should be appended to bottom of email template 
        Public AppendMarketingContent As Boolean



        Public CustomerCase As sCase

        Public Survey As sSurvey

        Public Customer As sCustomerEmail

        Public Registration As sRegistration





        Public IncludeRedirectReplyLink As Boolean
        Public CaseID As Integer
        Public CustomerID As Integer

    End Structure


    Public Structure sSurvey
        Dim Name As String
        Dim SurveyKey As String

    End Structure

    Public Structure sPop

        Dim ShoppingCartID As Integer

        Dim CustomerID As Integer

    End Structure






    '-------------------------------------------------------------------------------------

    Public Structure sCustomerEmail

        Dim CustomerID As Integer

        Dim CustomerName As String

        Dim Email As String

        Dim Address As String
        Dim IsBusiness As Boolean
        Dim Active As Boolean
        Dim Password As String
        Dim AccountStatus As AccountStatus
        Dim PasswordKey As String
        Dim CreditApplicationStatus As CreditApplicationStatus
        Dim TaxExemptCertificatePath As String
        Dim TaxExemptNumber As String
        Dim Comments As String





        Public Sub ResetPassword()

            Dim lngUserID As Long = 0 'Normalize Return Value

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLDR As SqlDataReader

            Dim SQLCommand As New SqlCommand

            Try

                '

                'SqlSelectCommand1

                '

                SQLCommand.CommandText = "CMS..[UpdateCustomerPassword]"

                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Password", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Password))

                SQLConn.Open()

                CustomerID = SQLCommand.ExecuteScalar





                If Not CustomerID > 0 Then

                    Throw New Exception("Error resetting password")

                End If

            Catch SQLErr As SqlException

                CustomerID = 0

                'TODO create sys message for role not found 

            Catch ex As Exception

                CustomerID = 0

                'TODO create sys message for role not found 

            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try

        End Sub

        Public Sub Load()

            Dim lngUserID As Long = 0 'Normalize Return Value

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLDR As SqlDataReader

            Dim SQLCommand As New SqlCommand

            Try

                '

                'SqlSelectCommand1

                '

                SQLCommand.CommandText = "LLFWebSitePrivate..[GetCustomerEmailInfo]"

                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))

                SQLConn.Open()

                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)



                If SQLDR.Read Then

                    CustomerName = SQLDR("Customer").ToString

                    CustomerID = SQLDR("CustomerID").ToString

                    Email = SQLDR("Email").ToString
                    IsBusiness = SQLDR("IsBusiness").ToString
                    Active = SQLDR("Active").ToString
                    CreditApplicationStatus = SQLDR("CreditApplicationStatusID").ToString
                    AccountStatus = SQLDR("AccountStatusID").ToString
                    TaxExemptCertificatePath = SQLDR("TaxExemptCertificatePath").ToString
                    TaxExemptNumber = SQLDR("TaxExemptNumber").ToString
                    Comments = SQLDR("Comments").ToString
                    'ProductID = SQLDR("ProductID")

                    'ProductModelNumber = SQLDR("ProductModelNumber")

                    'CustomerOwnershipID = SQLDR("CustomerOwnershipID")

                End If

            Catch SQLErr As SqlException

                'TODO create sys message for role not found 

            Catch ex As Exception

                'TODO create sys message for role not found 

            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try

        End Sub

        Public Sub LoadbyEmail()

            Dim lngUserID As Long = 0 'Normalize Return Value

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLDR As SqlDataReader

            Dim SQLCommand As New SqlCommand

            Try

                '

                'SqlSelectCommand1

                '

                SQLCommand.CommandText = "LLFWebSitePrivate..[GetCustomerAccountStatusInfobyEmail]"

                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Email", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Email))

                SQLConn.Open()

                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)



                If SQLDR.Read Then

                    CustomerName = SQLDR("Customer").ToString

                    CustomerID = SQLDR("CustomerID").ToString

                    Email = SQLDR("Email").ToString
                    Comments = SQLDR("Comments").ToString

                    Try
                        AccountStatus = SQLDR("AccountStatusID")
                    Catch ex As Exception
                        AccountStatus = BPS.AccountStatus.Unknown
                    End Try

                    'ProductName = SQLDR("ProductName")

                    'ProductID = SQLDR("ProductID")

                    'ProductModelNumber = SQLDR("ProductModelNumber")

                    'CustomerOwnershipID = SQLDR("CustomerOwnershipID")

                End If

            Catch SQLErr As SqlException

                'TODO create sys message for role not found 

            Catch ex As Exception

                'TODO create sys message for role not found 

            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try

        End Sub

    End Structure





    Public Structure sCase

        Dim CaseID As Integer

        Dim Description As String

        Dim CustomerID As Integer

        Dim CustomerName As String

        Dim FirstName As String

        Dim LastName As String

        Dim ProductName As String

        Dim ProductID As String

        Dim ProductModelNumber As String

        Dim CustomerOwnershipID As Integer

        Dim LatestUpdate As String

        Dim Resolution As String
        Dim IsBusiness As Boolean

        Dim Origin As String
        Dim bAccountRegistrated As Boolean

        Dim RoutingTypeID As Integer
        Dim CaseTypeID As Integer
        Dim CaseTypeName As String
        Dim RoutingBinName As String
        Dim CorrespondenceID As Integer
        Dim CorrespondenceBody As String
        Dim CorrespondenceSubject As String
        Dim IncludeCaseAttachment As Boolean
        Dim IncludeAppliedKnowlege As Boolean


        Dim DateCreated As String
        Dim DateLastUpdated As String
        Dim DateClosed As String
        Dim CaseStatus As String

        Dim AppliedKnowledgeLinks As String
        Dim ProductKnowledgeLinks As String
        Dim CaseAttachmentLinks As String

        Dim RoutingRecipients As String


        Dim Email As String

        Dim LastDateUpdated As String

       

        Public Sub Load()

            Dim lngUserID As Long = 0 'Normalize Return Value

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLDR As SqlDataReader

            Dim SQLCommand As New SqlCommand

            Try

                '

                'SqlSelectCommand1

                '

                SQLCommand.CommandText = "LLFWebSitePrivate..[GetNewCaseEmailInfo]"

                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CaseID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CaseID))

                SQLConn.Open()

                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)

                If SQLDR.Read Then

                    CustomerName = SQLDR("Customer")

                    LastName = SQLDR("LastName")

                    CustomerID = SQLDR("CustomerID")

                    Email = SQLDR("Email")

                    ProductName = SQLDR("ProductName")

                    ProductID = SQLDR("ProductID")

                    ProductModelNumber = SQLDR("ProductModelNumber")
                    Description = SQLDR("Description")
                    CaseStatus = SQLDR("Status")
                    CustomerOwnershipID = SQLDR("CustomerOwnershipID")

                    Origin = SQLDR("Origin")

                    bAccountRegistrated = SQLDR("IsRegisteredAccount")

                    ' SerialNumber = SQLDR("SerialNumber")

                    ' PurchaseDate = SQLDR("DateOfPurchase")

                    ' RetailerID = SQLDR("RetailerID")

                    ' RetailerName = SQLDR("RetailerName")

                    ' WarrantyClaimID = SQLDR("WarrantyClaimID")

                End If

            
            Catch SQLErr As SqlException

                'TODO create sys message for role not found 

            Catch ex As Exception

                'TODO create sys message for role not found 

            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try

        End Sub

    End Structure





    Public Structure sRegistration

        Dim CustomerOwnerShipID As Integer

        Dim CustomerName As String

        Dim FirstName As String

        Dim LastName As String

        Dim Email As String

        Dim POPRequestedCaseID As Integer

        Dim ProductName As String

        Dim ProductModelNumber As String

        Dim SerialNumber As String

        Dim ProductModelYear As String

        Dim ProductID As String

        Dim CustomerID As Integer

        Dim RegistrationDate As String

        Dim PurchaseDate As String

        Dim RetailerID As Integer

        Dim RetailerName As String

        Dim WarrantyClaimID As String

        Dim POPResolve As Integer





        Public Sub Load()

            Dim lngUserID As Long = 0 'Normalize Return Value

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("CMSBaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLDR As SqlDataReader

            Dim SQLCommand As New SqlCommand

            Try

                '

                'SqlSelectCommand1

                '

                SQLCommand.CommandText = "CMS..[GetCustomerRegistrationEmailInfo]"

                SQLCommand.CommandType = System.Data.CommandType.StoredProcedure

                SQLCommand.Connection = SQLConn

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerOwnershipID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerOwnerShipID))

                SQLConn.Open()

                SQLDR = SQLCommand.ExecuteReader(CommandBehavior.SingleRow)



                If SQLDR.Read Then

                    CustomerName = SQLDR("Customer")

                    LastName = SQLDR("LastName")

                    CustomerID = SQLDR("CustomerID")

                    Email = SQLDR("Email")

                    ProductName = SQLDR("ProductName")

                    ProductID = SQLDR("ProductID")

                    ProductModelNumber = SQLDR("ProductModelNumber")

                    SerialNumber = SQLDR("SerialNumber")

                    PurchaseDate = SQLDR("DateOfPurchase")

                    RetailerID = SQLDR("RetailerID")

                    RetailerName = SQLDR("RetailerName")

                    WarrantyClaimID = SQLDR("WarrantyClaimID")

                    POPRequestedCaseID = SQLDR("POPRequestedCaseID")

                End If

            Catch SQLErr As SqlException

                'TODO create sys message for role not found 

            Catch ex As Exception

                'TODO create sys message for role not found 

            Finally

                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()

                SQLConn = Nothing

                SQLDR = Nothing

                SQLCommand = Nothing

            End Try

        End Sub

    End Structure


    'Public Structure sEmail
    '    '"sEmail" Structure Profile ******************************************************
    '    '
    '    '       Structure:   sEmail
    '    '       Parent File: Structures.vb
    '    '       Parent Project: Business Logic
    '    '       Parent Solution: HLGIT CRM System
    '    '       Created By: Vincent Clover
    '    '       Created on: 07/02/04
    '    '       Description:  Can Hold information of a Email
    '    '
    '    '
    '    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '    Public MailToAddress As String
    '    Public MailFromAddress As String
    '    Public Subject As String
    '    Public Body As String
    '    Public SMTPServer As String
    '    Public Priority As EmailPriorities
    '    Public BodyFormat As EmailBodyFormat
    '    Public IsBodyHtml As Boolean
    '    Public ReplyTo As String
    '    Public Attachments As Collection


    '    Public CaseID As Integer
    '    Public CustomerID As Integer
    '    Public IncludeRedirectReplyLink As Boolean
    'End Structure
End Namespace
