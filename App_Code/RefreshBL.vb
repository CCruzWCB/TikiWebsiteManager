Imports System.Data
Imports System.Text
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Configuration

Namespace Refresh
    Public Class Product

        '"Product" Class Profile ******************************************************
        '
        '       Class: Product 
        '       Parent File: WorkElements.vb
        '       Parent Project: CMS
        '       Parent Solution: CMS
        '       Created By: Amy Cook
        '       Created on: 07/21/05
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _lngProductID As Long          'Used to Hold the ProductID currently being Used/Created
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs
#End Region

#Region "Class Properties"


        'Gets & Sets Current ProductID
        Public Property ProductID()
            Get
                Return _lngProductID  'Gets Local _lngProductID Variable
            End Get
            Set(ByVal Value)
                _lngProductID = Value 'Sets Local _lngProductID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage()
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber()
            Get
                Return _intErrorNumber
            End Get
        End Property


#End Region

#Region "Constructors"


        'Empty Consructor
        Public Sub New()

        End Sub

        'Use this Constructor when wanting to set the ProductID on Creation
        Public Sub New(ByVal ProductID As Integer)

            _lngProductID = ProductID 'Sets Local ProductID Variable

        End Sub






#End Region

#Region "Functions & Procedures"
        Public Function AddProduct(ByRef sProduct As sProduct) As Boolean

            'Function: AddProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProduct
            '       Parent Class:  Product
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProduct Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/16/05
            '       Description: Adds a Product 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "BPSv2..[AddProduct]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                With sProduct

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Name))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 1000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Description))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Freight", System.Data.SqlDbType.VarChar, 500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Freight))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSupported", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsSupported))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Specifications", System.Data.SqlDbType.VarChar, 1250, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Specifications))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Year", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Year))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_large", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDLarge))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_small", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDSmall))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_feature", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDFeature))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedBy))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedByName))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsActive))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductNumber))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .MSRP))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField1", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField1))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField2", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField2))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField3", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField3))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField4", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField4))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField5", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField5))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UPCCode", System.Data.SqlDbType.VarChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UPCCode))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Weight", System.Data.SqlDbType.VarChar, 10, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Weight))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AvailableQuantity", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AvailableQuantity))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ExpectedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ExpectedDate))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Surcharge", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Surcharge))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSellable", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsSellable))

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CompanyID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .BrandID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductSeriesID))


                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))



                    If sProduct.MSRP > 0 Then
                        SQLCommand.Parameters("@MSRP").Value = sProduct.MSRP
                    Else
                        SQLCommand.Parameters("@MSRP").Value = DBNull.Value
                    End If

                End With

                SQLCommand.ExecuteNonQuery()


                'Pass the new "KnowledgeID" back to the sKnowledge Structure
                sProduct.ProductID = SQLCommand.Parameters("@ProductID").Value

                If Not SQLCommand.Parameters("@ProductID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False
                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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




        Public Function GetProduct(ByRef sProduct As sProduct) As Boolean

            'Function: GetProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetProduct
            '       Parent Class:  Product
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProduct Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  11/21/05
            '       Description: Gets a Product 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


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

                SQLCommand.CommandText = "BPSv2..[GetProduct]"

                'Set the ProdcutID Parameter (IF Passed)
                If sProduct.ProductID > 0 Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProduct.ProductID))

                Else
                    If sProduct.ProductNumber <> "" Then
                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductNumber", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProduct.ProductNumber))
                    End If
                End If


                'Check to see if the BrandID was passed
                If sProduct.BrandID > 0 Then 'GET BY BRAND & Product
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProduct.BrandID))
                End If

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sProduct
                        'Fill Knowledge Information

                        If Not IsDBNull(SQLDataReader("ProductID")) Then .ProductID = SQLDataReader("ProductID")
                        If Not IsDBNull(SQLDataReader("ProductJunctionID")) Then .ProductJunctionID = SQLDataReader("ProductJunctionID")
                        If Not IsDBNull(SQLDataReader("ProductNumber")) Then .ProductNumber = SQLDataReader("ProductModelNumber")
                        If Not IsDBNull(SQLDataReader("ProductSeriesID")) Then .ProductSeriesID = SQLDataReader("ProductSeriesID")
                        If Not IsDBNull(SQLDataReader("SeriesName")) Then .ProductSeriesName = SQLDataReader("SeriesName")
                        If Not IsDBNull(SQLDataReader("Year")) Then .ProductYear = SQLDataReader("Year")
                        If Not IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("Freight")) Then .Freight = SQLDataReader("Freight")
                        If Not IsDBNull(SQLDataReader("IsSupported")) Then .IsSupported = SQLDataReader("IsSupported")
                        If Not IsDBNull(SQLDataReader("Specifications")) Then .Specifications = SQLDataReader("Specifications")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
                        If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        If Not IsDBNull(SQLDataReader("ImagePath_large")) Then .ImagePathLarge = SQLDataReader("ImagePath_large")
                        If Not IsDBNull(SQLDataReader("ImagePath_small")) Then .ImagePathSmall = SQLDataReader("ImagePath_small")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("Active")) Then .IsActive = SQLDataReader("Active")
                        If Not IsDBNull(SQLDataReader("CompanyID")) Then .CompanyID = SQLDataReader("CompanyID")
                        If Not IsDBNull(SQLDataReader("CompanyName")) Then .CompanyName = SQLDataReader("CompanyName")
                        If Not IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
                        If Not IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
                        'If Not IsDBNull(SQLDataReader("FirstCategoryID")) Then .ProductCategory1 = SQLDataReader("FirstCategoryID")
                        'If Not IsDBNull(SQLDataReader("FirstCategoryName")) Then .ProductCategory1Name = SQLDataReader("FirstCategoryName")
                        'If Not IsDBNull(SQLDataReader("SecondCategoryID")) Then .ProductCategory2 = SQLDataReader("SecondCategoryID")
                        'If Not IsDBNull(SQLDataReader("SecondCategoryName")) Then .ProductCategory2Name = SQLDataReader("SecondCategoryName")
                        If Not IsDBNull(SQLDataReader("MSRP")) Then .MSRP = SQLDataReader("MSRP")
                        If Not IsDBNull(SQLDataReader("IsSellable")) Then .IsSellable = SQLDataReader("IsSellable")
                        If Not IsDBNull(SQLDataReader("AvailableQuantity")) Then .AvailableQuantity = SQLDataReader("AvailableQuantity")
                        If Not IsDBNull(SQLDataReader("ExpectedDate")) Then .ExpectedDate = SQLDataReader("ExpectedDate")

                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Product Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If
                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess


        End Function

        Public Function GetProductID(ByRef ProductID As Long, ByVal ProductNumber As String, Optional ByVal BrandID As Integer = 0) As Boolean

            'Function: GetProductID ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetProductID
            '       Parent Class:  Product
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductNumber
            '       Optional Parameters(1): BrandID
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/23/07
            '       Description: Gets a ProductID
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 
                ProductID = 0 'Preset

                SQLConn.Open() 'Open Database

                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                          'Set the Connection

                SQLCommand.CommandText = "BPSv2..[GetProduct]"



                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductNumber", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ProductNumber))


                'Check to see if the BrandID was passed
                If BrandID > 0 Then 'GET BY BRAND & Product
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, BrandID))
                End If

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    ProductID = SQLDataReader("ProductID")
                    bSuccess = True


                Else
                    bSuccess = False
                    _strErrorMessage = "Product Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If
                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess


        End Function





        Public Function GetProduct(ByRef sProduct As sProduct, ByVal ProductNumber As String, Optional ByVal CompanyID As Long = 0, Optional ByVal BrandID As Long = 0) As Boolean

            'Function: GetProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetProduct
            '       Parent Class:  Product
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProduct Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/20/06
            '       Description: Gets a Product 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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

                SQLCommand.CommandText = "BPSv2..[GetProductByProductNumber]"

                'Set the Product Model Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductNumber", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ProductNumber))

                If CompanyID > 0 Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, CompanyID))
                End If

                If BrandID > 0 Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, BrandID))
                End If



                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)

                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sProduct
                        'Fill Knowledge Information
                        If Not IsDBNull(SQLDataReader("ProductID")) Then .ProductID = SQLDataReader("ProductID")
                        If Not IsDBNull(SQLDataReader("ProductJunctionID")) Then .ProductJunctionID = SQLDataReader("ProductJunctionID")
                        If Not IsDBNull(SQLDataReader("ProductNumber")) Then .ProductNumber = SQLDataReader("ProductNumber")
                        If Not IsDBNull(SQLDataReader("ProductSeriesID")) Then .ProductSeriesID = SQLDataReader("ProductSeriesID")
                        If Not IsDBNull(SQLDataReader("ProductSeriesName")) Then .ProductSeriesName = SQLDataReader("ProductSeriesName")
                        If Not IsDBNull(SQLDataReader("ProductName")) Then .Name = SQLDataReader("ProductName")
                        If Not IsDBNull(SQLDataReader("ProductDescription")) Then .Description = SQLDataReader("ProductDescription")
                        If Not IsDBNull(SQLDataReader("Freight")) Then .Freight = SQLDataReader("Freight")
                        If Not IsDBNull(SQLDataReader("IsSupported")) Then .IsSupported = SQLDataReader("IsSupported")
                        If Not IsDBNull(SQLDataReader("Specifications")) Then .Specifications = SQLDataReader("Specifications")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
                        If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("Active")) Then .IsActive = SQLDataReader("Active")
                        If Not IsDBNull(SQLDataReader("CompanyID")) Then .CompanyID = SQLDataReader("CompanyID")
                        If Not IsDBNull(SQLDataReader("CompanyName")) Then .CompanyName = SQLDataReader("CompanyName")
                        If Not IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
                        If Not IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
                        ' If Not IsDBNull(SQLDataReader("FirstCategoryID")) Then .ProductCategory1 = SQLDataReader("FirstCategoryID")
                        ' If Not IsDBNull(SQLDataReader("FirstCategoryName")) Then .ProductCategory1Name = SQLDataReader("FirstCategoryName")
                        ' If Not IsDBNull(SQLDataReader("SecondCategoryID")) Then .ProductCategory2 = SQLDataReader("SecondCategoryID")
                        ' If Not IsDBNull(SQLDataReader("SecondCategoryName")) Then .ProductCategory2Name = SQLDataReader("SecondCategoryName")
                        If Not IsDBNull(SQLDataReader("ProductMSRP")) Then .MSRP = SQLDataReader("ProductMSRP")
                        If Not IsDBNull(SQLDataReader("IsSellable")) Then .IsSellable = SQLDataReader("IsSellable")
                        If Not IsDBNull(SQLDataReader("AvailableQuantity")) Then .AvailableQuantity = SQLDataReader("AvailableQuantity")
                        If Not IsDBNull(SQLDataReader("ExpectedDate")) Then .ExpectedDate = SQLDataReader("ExpectedDate")


                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Product Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If
                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function





#End Region
    End Class

    Public Class ProductSeries

        '"Product" Class Profile ******************************************************
        '
        '       Class: ProductSeries 
        '       Parent File: WorkElements.vb
        '       Parent Project: CharbroilWarranty
        '       Parent Solution: CharbroilWarranty
        '       Created By: Amy Cook
        '       Created on: 11/10/05
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _lngProductSeriesID As Long          'Used to Hold the ProductSeriesID currently being Used/Created
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs
#End Region

#Region "Class Properties"


        'Gets & Sets Current ProductID
        Public Property ProductSeriesID()
            Get
                Return _lngProductSeriesID  'Gets Local _lngProductSeriesID Variable
            End Get
            Set(ByVal Value)
                _lngProductSeriesID = Value 'Sets Local _lngProductID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage()
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber()
            Get
                Return _intErrorNumber
            End Get
        End Property


#End Region

#Region "Constructors"


        'Empty Consructor
        Public Sub New()

        End Sub

        'Use this Constructor when wanting to set the ProductID on Creation
        Public Sub New(ByVal ProductSeriesID As Integer)

            _lngProductSeriesID = ProductSeriesID 'Sets Local ProductID Variable

        End Sub




        ' Use this Constructor when wanting to create a NEW ProductSeries(Requires a "sProductSeries") 
        Public Sub New(ByRef sProductSeries As sProductSeries, ByRef CreatedProductSeries As Boolean)

            CreatedProductSeries = AddProductSeries(sProductSeries)

        End Sub




#End Region

#Region "Functions & Procedures"

        Public Function AddProductSeries(ByRef sProductSeries As sProductSeries) As Boolean

            'Function: AddProductSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductSeries
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductSeries Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/16/05
            '       Description: Adds a Product Series
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_AddProductSeries]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure

                With sProductSeries

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Name))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSellable", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsSellable))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 8000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Description))
                    ' SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WebDescription", System.Data.SqlDbType.VarChar, 4000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .WebDescription))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Freight", System.Data.SqlDbType.VarChar, 500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Freight))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSupported", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsSupported))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Specifications", System.Data.SqlDbType.VarChar, 1250, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Specifications))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_large", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDLarge))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_small", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDSmall))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_feature", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDFeature))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedBy))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedByName))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Active))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WarrantyInfo", System.Data.SqlDbType.VarChar, 1500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .WarrantyInfo))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductSeriesNumber))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsBundle", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsBundle))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .MSRP))
                    '
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Bonus", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Bonus))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Multiplier", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Multiplier))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@HasSubProducts", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .HasSubProducts))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsOversized", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsOversized))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField1", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField1))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField2", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField2))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsPersonalized", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsPersonalized))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PersonalizationDescription", System.Data.SqlDbType.VarChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PersonalizationDescription))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CompanyID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .BrandID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductCategoryID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AvailableQuantity", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AvailableQuantity))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ExpectedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ExpectedDate))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AdditionalShipping", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AdditionalShipping))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductGroupName", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductGroupName))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AttributeTypeID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AttributeTypeID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AttributeTypeName", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AttributeTypeName))


                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                End With
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                If Not SQLCommand.Parameters("@ProductSeriesID").Value > 0 OrElse Not SQLCommand.Parameters("@ProductSeriesID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product Series(Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else
                    'Pass the new "KnowledgeID" back to the sKnowledge Structure
                    sProductSeries.ProductSeriesID = SQLCommand.Parameters("@ProductSeriesID").Value
                    sProductSeries.ProductID = SQLCommand.Parameters("@ProductID").Value

                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function AddProductBullets(ByVal ProductID As Integer, ByVal BulletList As String) As Boolean

            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series Attribute One Step is Required: ********************
            '   - 1. Add the Product Series Attribute table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_AddProductBullets]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AttributeTypeID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductSeriesAttributeType.Bullet))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BulletList", System.Data.SqlDbType.VarChar, 5000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, BulletList))


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function AddProductRetailers(ByVal ProductID As Integer, ByVal RetailerList As String) As Boolean

            'Function: AddProductSeriesAttribute ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductSeriesAttribute
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductSeriesAttribute Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/16/07
            '       Description: Adds a Product Series Attribute
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series Attribute One Step is Required: ********************
            '   - 1. Add the Product Series Attribute table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_AddProductRetailers]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure



                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RetailerList", System.Data.SqlDbType.VarChar, 900, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, RetailerList))


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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




        Public Function AddProductSeriesAttribute(ByRef sProductSeriesAttribute As sProductSeriesAttribute) As Boolean

            'Function: AddProductSeriesAttribute ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductSeriesAttribute
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductSeriesAttribute Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/16/07
            '       Description: Adds a Product Series Attribute
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series Attribute One Step is Required: ********************
            '   - 1. Add the Product Series Attribute table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_AddProductAttribute]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure

                With sProductSeriesAttribute

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductSeriesID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AttributeTypeID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AttributeTypeID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AttributeName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AttributeName))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AttributeValue", System.Data.SqlDbType.VarChar, 250, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AttributeValue))


                End With
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function AddProductToRepPortal(ByRef sProductSeries As sProductSeries) As Boolean


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
                SQLCommand.CommandText = "LLRepPortal.Refresh.[AddProduct]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure

                With sProductSeries

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductModelNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductSeriesNumber))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Name))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 8000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Description))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CaseQty", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CaseQty))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Active))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsOversized", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsOversized))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsHidden", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsHidden))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsCaseQtyOnly", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsCaseQTYOnly))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                End With

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                If Not SQLCommand.Parameters("@ProductID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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



        Public Function AddProductAssociation(ByVal ProductID As Long, ByVal AssociatedProductID As Long, ByVal ProductSeriesAssociatedType As ProductSeriesAssociatedType) As Boolean


            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series Attribute One Step is Required: ********************
            '   - 1. Add the Product Series Attribute table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_AddProductAssociation]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure



                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AssociatedProductID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, AssociatedProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AssociatedTypeID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductSeriesAssociatedType))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AssociatedProductsID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function DeleteProductSeriesAssociations(ByVal ProductSeriesID As Long) As Boolean

            'Function: DeleteProductSeriesAssociations ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteProductSeriesAssociations
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductSeriesID 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/17/07
            '       Description: Removes All Product Series Association by ProductSeries
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Remove Product Series Associations One Step is Required: ********************
            '   - 1. Add the Product Series Attribute table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "BPSv2..[DeleteProductSeriesAssociations]"     'Stored Procedure Name

                'Stored Procedure Paramaters - 
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductSeriesID))


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                'Close Connection
                SQLConn.Close()
                bSuccess = True

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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





        Public Function DeleteProductSeriesJUNC(Optional ByVal ProductSeriesJunctionID As Long = 0, Optional ByVal ProductSeriesID As Long = 0) As Boolean


            'Function:  DeleteProductSeriesJUNC ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name:  DeleteProductSeriesJUNC
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductSeriesJunctionID
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/20/07
            '       Description: Deletes a Product Series JUNC entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[DeleteProductSeriesJunc]"     'Stored Procedure Name

                If ProductSeriesJunctionID > 0 Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesJunctionID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductSeriesJunctionID))

                Else
                    If ProductSeriesID > 0 Then
                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductSeriesID))
                    End If
                End If

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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




        Public Function AddProductGroups(ByVal BrandID As Integer) As Boolean


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
                SQLCommand.CommandText = "LLFBPS..[r_AddGroupProducts]"     'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, BrandID))
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function DeleteAllAssociatedProducts(ByVal BrandID As Integer) As Boolean


            'Function:  DeleteProductSeriesAttributes ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name:  DeleteProductSeriesAttributes
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): BrandID
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/20/07
            '       Description: Deletes a Product Series Attributes
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "LLFBPS..[r_DeleteAllAssociatedProducts]"     'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, BrandID))
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function DeleteProductAttributes(ByVal ProductID As Long) As Boolean


            'Function:  DeleteProductSeriesAttributes ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name:  DeleteProductSeriesAttributes
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductSeriesID
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/20/07
            '       Description: Deletes a Product Series Attributes
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "LLFBPS..[r_DeleteProductAttributes]"     'Stored Procedure Name

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductID))

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function DisableAllProducts(ByVal ProductSeriesID As Long) As Boolean


            'Function:  DisableAllProducts ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name:  DisableAllProducts
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductSeriesID
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/20/07
            '       Description: Disables All Products by ProductSeriesID
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[DisableAllProductsByProductSeries]"     'Stored Procedure Name


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductSeriesID))


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function AddKeyCodePrice(ByVal ItemNumber As String, ByVal PriceListID As Integer, ByVal MSRP As Double) As Boolean


            'Function: ResetAllKeyCodePrices ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name:  ResetAllKeyCodePrices
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): PriceListID
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  01/9/09
            '       Description: 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "LLFBPS..[r_AddKeyCodePrice]"     'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Itemnumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ItemNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PriceListID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, PriceListID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, MSRP))
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function ResetAllKeyCodePrices(ByVal PriceListID As Integer) As Boolean


            'Function: ResetAllKeyCodePrices ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name:  ResetAllKeyCodePrices
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): PriceListID
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  01/9/09
            '       Description: 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "LLFBPS..[r_ResetAllKeyCodePrices]"     'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PriceListID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, PriceListID))
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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




        Public Function ResetAllEcomProducts(ByVal BrandID As Integer) As Boolean


            'Function:  DisableAllProducts ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name:  DisableAllProducts
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductSeriesID
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/20/07
            '       Description: Disables All Products by ProductSeriesID
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "LLFBPS..[r_ResetAllEcomProducts]"     'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, BrandID))
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function ResetAllRepProducts() As Boolean


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
                SQLCommand.CommandText = "LLRepPortal.Refresh.[DeleteAllProducts]"     'Stored Procedure Name
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function DisableEmptyEcomCategoryies(ByVal BrandID As Integer) As Boolean


            'Function:  DisableEmptyEcomCategoryies ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name:  DisableEmptyEcomCategoryies
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductSeriesID
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/20/07
            '       Description: Disables All Products by ProductSeriesID
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "LLFBPS..[r_DisableEmptyEcomCategoryies]"     'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, BrandID))
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function AddProductSeriesJUNC(ByRef sProductSeries As sProductSeries) As Boolean

            'Function: AddProductSeriesJUNC ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductSeriesJUNC
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductSeries Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/20/07
            '       Description: Adds a Product Series JUNC
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series JUNC table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "BPSv2..[AddProductSeriesJunc]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                With sProductSeries

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductSeriesID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CompanyID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .BrandID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductCategoryID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedBy))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedByName))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesJunctionID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))

                End With

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "KnowledgeID" back to the sKnowledge Structure
                sProductSeries.ProductSeriesJunctionID = SQLCommand.Parameters("@ProductSeriesJunctionID").Value

                If Not SQLCommand.Parameters("@ProductSeriesJunctionID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product Series Relationship(Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number

                Select Case SQLErr.Number
                    Case 2627
                        _strErrorMessage = "This series already belongs to this category." 'Set the Classes Error Message
                    Case Else
                        _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                End Select
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function GetProductSeries(ByRef sProductSeries As sProductSeries) As Boolean

            'Function: GetProductSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetProductSeries
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductSeries Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/14/05
            '       Description: Gets a Product Series
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            'If sProductSeries.ProductSeriesID > 0 Then
            '    _lngProductSeriesID = sProductSeries.ProductSeriesID
            'End If

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

                SQLCommand.CommandText = "LLFBPS..[r_GetProductSeries]"



                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesNumber", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProductSeries.ProductSeriesNumber))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sProductSeries
                        'Fill Knowledge Information
                        If Not IsDBNull(SQLDataReader("OriginalProductSeriesID")) Then .ProductSeriesID = SQLDataReader("OriginalProductSeriesID")
                        If Not IsDBNull(SQLDataReader("ProductID")) Then .ProductID = SQLDataReader("ProductID")

                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Product Series Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If
                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function

        Public Function UpdateProductSeries(ByRef sProductSeries As sProductSeries) As Boolean

            'Function: UpdateProductSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateProductSeries
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductSeries Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/16/05
            '       Description: Updates a Product Series
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Update Product Series One Step is Required: ********************
            '   - 1. Update the ProductSeries table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection

                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_UpdateProductSeries]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure

                With sProductSeries

                    Dim e As String
                    e = "e"


                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductSeriesID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Name))

                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WebDescription", System.Data.SqlDbType.VarChar, 4000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .WebDescription))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Freight", System.Data.SqlDbType.VarChar, 500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Freight))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSellable", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsSellable))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Specifications", System.Data.SqlDbType.VarChar, 1250, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Specifications))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_large", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDLarge))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_small", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDSmall))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_feature", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageIDFeature))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedBy))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedByName))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Active))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WarrantyInfo", System.Data.SqlDbType.VarChar, 1500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .WarrantyInfo))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductSeriesNumber))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsBundle", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsBundle))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .MSRP))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AdditionalShipping", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AdditionalShipping))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Bonus", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Bonus))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Multiplier", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Multiplier))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@HasSubProducts", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .HasSubProducts))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsOversized", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsOversized))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField1", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField1))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField2", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .UserField2))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsPersonalized", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .IsPersonalized))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PersonalizationDescription", System.Data.SqlDbType.VarChar, 200, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PersonalizationDescription))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CompanyID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .BrandID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductCategoryID))
                    'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesJunctionID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductSeriesJunctionID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AvailableQuantity", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AvailableQuantity))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ExpectedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ExpectedDate))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AdditionalShipping", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AdditionalShipping))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 8000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Description))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductGroupName", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductGroupName))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AttributeTypeID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AttributeTypeID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AttributeTypeName", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .AttributeTypeName))

                End With

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                ' If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                bSuccess = True
                'Else
                'bSuccess = False

                'End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
                bSuccess = False

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLConn = Nothing
                SQLCommand = Nothing

            End Try

            'Set returning Boolean
            ' bSuccess = True
            Return bSuccess

        End Function


#End Region
    End Class






    Public Class Part

        '"Part" Class Profile ******************************************************
        '
        '       Class: Part 
        '       Parent File: WorkElements.vb
        '       Parent Project: BPS
        '       Parent Solution: BPS
        '       Created By: Amy Cook
        '       Created on: 12/05/05
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _lngPartID As Long          'Used to Hold the PartID currently being Used/Created
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs
#End Region

#Region "Class Properties"


        'Gets & Sets Current ProductID
        Public Property PartID()
            Get
                Return _lngPartID  'Gets Local _lngProductID Variable
            End Get
            Set(ByVal Value)
                _lngPartID = Value 'Sets Local _lngProductID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage()
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber()
            Get
                Return _intErrorNumber
            End Get
        End Property


#End Region

#Region "Constructors"


        'Empty Consructor
        Public Sub New()

        End Sub

        'Use this Constructor when wanting to set the PartID on Creation
        Public Sub New(ByVal PartID As Integer)

            _lngPartID = PartID 'Sets Local PartID Variable

        End Sub




        ' Use this Constructor when wanting to create a NEW Product(Requires a "sPart") 
        Public Sub New(ByRef sPart As sPart, ByRef CreatedPart As Boolean)

            CreatedPart = AddPart(sPart)

        End Sub




#End Region

#Region "Functions & Procedures"
        Public Function AddPartAlias(ByVal PartNumber As String, ByVal PartAlias As String)

            'Function: AddPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddPartAlias
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): PartID
            '       Required Parameters(2): PartAlias
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/03/06
            '       Description: Adds a Part Alias
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Part - One Step is Required: ********************
            '   - 1. Add the Part table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "BPSv2..[AddPartAlias]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sPart" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartAlias", System.Data.SqlDbType.VarChar, 75))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartAliasID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))

                SQLCommand.Parameters("@PartNumber").Value = PartNumber
                SQLCommand.Parameters("@PartAlias").Value = PartAlias
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                If Not SQLCommand.Parameters("@PartAliasID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Part Alias (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function AddPart(ByRef sPart As sPart) As Boolean

            'Function: AddPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddPart
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: zBPS
            '       Parent Solution: zBPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  6/09/08
            '       Description: Adds a Part
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Part - One Step is Required: ********************
            '   - 1. Add the Part table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_AddPart]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sPart" Structure



                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSellable", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))





                SQLCommand.Parameters("@MSRP").Value = sPart.MSRP
                SQLCommand.Parameters("@IsSellable").Value = sPart.IsSellable
                SQLCommand.Parameters("@Name").Value = sPart.Name
                SQLCommand.Parameters("@Description").Value = sPart.Description
                SQLCommand.Parameters("@PartNumber").Value = sPart.PartNumber
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "PartID" back to the sPart Structure
                sPart.PartID = SQLCommand.Parameters("@PartID").Value

                If Not SQLCommand.Parameters("@PartID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Part (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function ApplyPartToProduct(ByVal PartID As Integer, ByVal ProductID As Integer)

            'Function: ApplyPartToProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: ApplyPartToProduct
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/05/05
            '       Description: Applies a Part to a Product
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return value
            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Part - One Step is Required: ********************
            '   - 1. Add the Part table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "BPSv2..[ApplyPartToProduct]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sPart" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@PartID").Value = PartID
                SQLCommand.Parameters("@ProductID").Value = ProductID
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function AddPartToProduct(ByRef sPart As sPart) As Boolean

            'Function: AddPartToProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddPartToProduct
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/05/05
            '       Description: Adds a Part and a PartProductJUNC
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return Value
            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Part - One Step is Required: ********************
            '   - 1. Add the Part table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "BPSv2..[AddPartToProduct]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sPart" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BaseWarrantyDuration", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CallLetter", System.Data.SqlDbType.Char, 5))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@Name").Value = sPart.Name
                SQLCommand.Parameters("@Description").Value = sPart.Description
                SQLCommand.Parameters("@PartNumber").Value = sPart.PartNumber
                SQLCommand.Parameters("@BaseWarrantyDuration").Value = sPart.BaseWarrantyDuration
                SQLCommand.Parameters("@CreatedBy").Value = sPart.CreatedBy
                SQLCommand.Parameters("@CreatedName").Value = sPart.CreatedByName
                SQLCommand.Parameters("@ProductID").Value = sPart.ProductID
                SQLCommand.Parameters("@PartCategoryID").Value = sPart.PartCategoryID
                SQLCommand.Parameters("@CallLetter").Value = sPart.CallLetter
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "PartID" back to the sPart Structure
                sPart.PartID = SQLCommand.Parameters("@PartID").Value

                If Not SQLCommand.Parameters("@PartID").Value > 0 Then
                    _strErrorMessage = "Internal Error adding Part (Not System Error)"     'Set the Classes Error Message
                    _intErrorNumber = 0     'Set the Classes Error Number
                    bSuccess = False
                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function AddPartToProductSeries_OLD(ByRef sPart As sPart) As Boolean

            'Function: AddPartToProductSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddPartToProductSeries
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/05/05
            '       Description: Adds a Part and a PartProductSeriesJUNC
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Part - One Step is Required: ********************
            '   - 1. Add the Part table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "BPSv2..[AddPartToProductSeries]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sPart" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@Name").Value = sPart.Name
                SQLCommand.Parameters("@PartNumber").Value = sPart.PartNumber
                SQLCommand.Parameters("@CreatedBy").Value = sPart.CreatedBy
                SQLCommand.Parameters("@CreatedName").Value = sPart.CreatedByName
                SQLCommand.Parameters("@ProductSeriesID").Value = sPart.ProductSeriesID
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "PartID" back to the sPart Structure
                sPart.PartID = SQLCommand.Parameters("@PartID").Value

                If Not SQLCommand.Parameters("@PartID").Value > 0 Then
                    _strErrorMessage = "Internal Error adding Part (Not System Error)"     'Set the Classes Error Message
                    _intErrorNumber = 0     'Set the Classes Error Number
                    bSuccess = False

                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function UpdatePartAliasPrices() As Boolean

            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)
            Dim DR As SqlDataReader
            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Dim dtParts As DataTable 'Customer Data Table 
            Dim drPart As DataRow   'Customer Data Row

            Try    'Only one "Try" statement 


                dtParts = New DataTable("Parts")
                dtParts.Columns.Add("PartNumber", GetType(String))


                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection



                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_GetAllNonSellableParts]"     'Stored Procedure Name
                DR = SQLCommand.ExecuteReader     'Execute the Stored Procedure



                Do While DR.Read
                    drPart = dtParts.NewRow
                    drPart(0) = DR(0)
                    dtParts.Rows.Add(drPart)

                Loop

                bSuccess = True

                'Close Connection
                DR.Close()



                'Now loop through each Non sellable part and update it by its alias

                SQLCommand.Parameters.Clear()
                SQLCommand.CommandText = "LLFBPS..[r_UpdatePartAliasPrice]"
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 40))


                For Each drPart In dtParts.Rows

                    Try



                        SQLCommand.Parameters("@PartNumber").Value = drPart(0)
                        SQLCommand.ExecuteNonQuery()

                    Catch ex As Exception

                    End Try

                Next


                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function ResetAllPartPriceFlags() As Boolean

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
                SQLCommand.CommandText = "LLFBPS..[r_ResetAllPartPriceFlags]"     'Stored Procedure Name
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function UpdatePartPrice(ByVal PartNumber As String, ByVal MSRP As Double) As Boolean

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
                SQLCommand.CommandText = "LLFBPS..[r_UpdatePartPrice]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@PartNumber").Value = PartNumber
                SQLCommand.Parameters("@MSRP").Value = MSRP


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function UpdateProductPartWarrantyDuration(ByVal sPart As sPart) As Boolean

            'Function: UpdateProductPartWarrantyDuration ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateProductPartWarrantyDuration
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/08/05
            '       Description: Updates a Product Part Warranty Duration
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[UpdateProductPartWarrantyDuration]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartProductID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WarrantyDuration", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@PartProductID").Value = sPart.PartProductID
                SQLCommand.Parameters("@WarrantyDuration").Value = sPart.PartProductWarrantyDuration

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Updating Part (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                End If


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function UpdateProductPart(ByVal sPart As sPart) As Boolean

            'Function: UpdateProductPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateProductPart
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/08/05
            '       Description: Updates a Product Part 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[UpdateProductPart]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BaseWarrantyDuration", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartProductID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WarrantyDuration", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CallLetter", System.Data.SqlDbType.Char, 5))

                SQLCommand.Parameters("@PartID").Value = sPart.PartID
                SQLCommand.Parameters("@Name").Value = sPart.Name
                SQLCommand.Parameters("@Description").Value = sPart.Description
                SQLCommand.Parameters("@PartNumber").Value = sPart.PartNumber
                SQLCommand.Parameters("@PrimaryResourceImageID").Value = sPart.PrimaryResourceImageID
                SQLCommand.Parameters("@BaseWarrantyDuration").Value = sPart.BaseWarrantyDuration
                SQLCommand.Parameters("@PartProductID").Value = sPart.PartProductID
                SQLCommand.Parameters("@Active").Value = sPart.Active
                SQLCommand.Parameters("@WarrantyDuration").Value = sPart.PartProductWarrantyDuration
                SQLCommand.Parameters("@PartCategoryID").Value = sPart.PartCategoryID
                SQLCommand.Parameters("@CallLetter").Value = sPart.CallLetter

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Updating Part (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                End If


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function UpdateProductSeriesPart_OLD(ByVal sPart As sPart) As Boolean

            'Function: UpdateProductSeriesPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateProductSeriesPart
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/08/05
            '       Description: Updates a Product Series Part 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[UpdateProductSeriesPart]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BaseWarrantyDuration", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WarrantyDuration", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@PartID").Value = sPart.PartID
                SQLCommand.Parameters("@Name").Value = sPart.Name
                SQLCommand.Parameters("@Description").Value = sPart.Description
                SQLCommand.Parameters("@PartNumber").Value = sPart.PartNumber
                SQLCommand.Parameters("@PrimaryResourceImageID").Value = sPart.PrimaryResourceImageID
                SQLCommand.Parameters("@BaseWarrantyDuration").Value = sPart.BaseWarrantyDuration
                SQLCommand.Parameters("@PartProductSeriesID").Value = sPart.PartProductSeriesID
                SQLCommand.Parameters("@Active").Value = sPart.Active
                SQLCommand.Parameters("@WarrantyDuration").Value = sPart.PartProductWarrantyDuration

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Updating Part (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                End If


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function DeleteProductSeriesPart(ByVal PartProductSeriesID As Long) As Boolean

            'Function: DeleteProductSeriesPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteProductSeriesPart
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): PartProductSeriesID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/17/05
            '       Description: Deletes a PartProductSeriesJUNC entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[DeletePartProductSeriesJUNC]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartProductSeriesID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@PartProductSeriesID").Value = PartProductSeriesID

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function DeletePartAlias(ByVal PartAliasID As Long) As Boolean

            'Function: DeletePartAlias ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeletePartAlias
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): PartAliasID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/17/05
            '       Description: Deletes a Part Alias entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[DeletePartAlias]"     'Stored Procedure Name

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartAliasID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@PartAliasID").Value = PartAliasID

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                Return True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function DeleteProductPart(ByVal PartProductID As Long) As Boolean

            'Function: DeleteProductPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteProductPart
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): PartProductID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/17/05
            '       Description: Deletes a PartProductSeriesJUNC entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[DeletePartProductJUNC]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartProductID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@PartProductID").Value = PartProductID

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                bSuccess = True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function GetProductPart(ByRef sPart As sPart) As Boolean

            'Function: GetProductPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetProductPart
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/05/05
            '       Description: Gets a Part
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If sPart.PartID > 0 Then
                _lngPartID = sPart.PartID
            End If

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

                SQLCommand.CommandText = "BPSv2..[GetProductParts]"

                'Set the ProdcutID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sPart.ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngPartID))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sPart
                        'Fill Part Information
                        If Not IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("PartNumber")) Then .PartNumber = SQLDataReader("PartNumber")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageName")) Then .PrimaryResourceImageName = SQLDataReader("PrimaryResourceImageName")
                        If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("BaseWarrantyDuration")) Then .BaseWarrantyDuration = SQLDataReader("BaseWarrantyDuration")
                        If Not IsDBNull(SQLDataReader("PartProductID")) Then .PartProductID = SQLDataReader("PartProductID")
                        If Not IsDBNull(SQLDataReader("ProductModelNumber")) Then .ProductModelNumber = SQLDataReader("ProductModelNumber")
                        If Not IsDBNull(SQLDataReader("ProductName")) Then .ProductName = SQLDataReader("ProductName")
                        If Not IsDBNull(SQLDataReader("WarrantyDuration")) Then .PartProductWarrantyDuration = SQLDataReader("WarrantyDuration")
                        If Not IsDBNull(SQLDataReader("PartCategoryID")) Then .PartCategoryID = SQLDataReader("PartCategoryID")
                        If Not IsDBNull(SQLDataReader("PartCategoryName")) Then .PartCategoryName = SQLDataReader("PartCategoryName")

                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Part Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If
                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            Return bSuccess

        End Function










        Public Function GetPart(ByRef sPart As sPart) As Boolean

            'Function: GetPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetPart
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/05/05
            '       Description: Gets a Part
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If sPart.PartID > 0 Then
                _lngPartID = sPart.PartID
            End If

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
                SQLCommand.CommandText = "BPSv2..[GetPart]"

                'Set the ProductID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngPartID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sPart.PartNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sPart.CompanyID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sPart.BrandID))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sPart
                        'Fill Part Information
                        If Not IsDBNull(SQLDataReader("PartID")) Then .PartID = SQLDataReader("PartID")
                        If Not IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("PartNumber")) Then .PartNumber = SQLDataReader("PartNumber")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageName")) Then .PrimaryResourceImageName = SQLDataReader("PrimaryResourceImageName")
                        If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("BaseWarrantyDuration")) Then .BaseWarrantyDuration = SQLDataReader("BaseWarrantyDuration")
                        If Not IsDBNull(SQLDataReader("PartCategoryID")) Then .PartCategoryID = SQLDataReader("PartCategoryID")
                        If Not IsDBNull(SQLDataReader("PartCategoryName")) Then .PartCategoryName = SQLDataReader("PartCategoryName")
                        If Not IsDBNull(SQLDataReader("MSRP")) Then .MSRP = SQLDataReader("MSRP")
                        If Not IsDBNull(SQLDataReader("IsSellable")) Then .IsSellable = SQLDataReader("IsSellable")

                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Part Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If
                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            Return bSuccess

        End Function



        Public Function GetPartAliasNumber(ByVal PartNumber As String, ByVal CompanyID As Integer) As String

            'Function: GetPartAliasNumber ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetPartAliasNumber
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: String
            '       Required Parameters(2): PartNumber) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  12/08/06
            '       Description: Gets a Part Alias Number
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            Dim ReturnPartNumber As String = PartNumber
            Dim blnContinue As Boolean = True
            Dim I As Integer = 1
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database

                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                          'Set the Connection
                SQLCommand.CommandText = "BPSv2..[GetPartAliasNumber]"


                Do Until blnContinue = False OrElse I = 24
                    blnContinue = False

                    SQLCommand.Parameters.Clear()
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, CompanyID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 75, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ReturnPartNumber))

                    'Set the DataReader
                    SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                    'Populate the Product Structure

                    If SQLDataReader.Read Then
                        I += 1
                        blnContinue = True
                        ReturnPartNumber = SQLDataReader("PartAlias")

                    End If



                    SQLDataReader.Close()
                Loop

                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            Return ReturnPartNumber

        End Function



#End Region
    End Class


    Public Class ProductCategory

        '"ProductCategory" Class Profile ******************************************************
        '
        '       Class: ProductCategory
        '       Parent File: WorkElements.vb
        '       Parent Project: CharbroilWarranty
        '       Parent Solution: CharbroilWarranty
        '       Created By: Vincent Clover
        '       Created on: 11/16/05
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _lngProductCategoryID As Long  'Used to Hold the ProductCategoryID currently being Used/Created
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs
#End Region

#Region "Class Properties"


        'Gets & Sets Current ProductCategoryID
        Public Property ProductCategoryID()
            Get
                Return _lngProductCategoryID  'Gets Local _lngProductCategoryID Variable
            End Get
            Set(ByVal Value)
                _lngProductCategoryID = Value 'Sets Local _lngProductCategoryID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage()
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber()
            Get
                Return _intErrorNumber
            End Get
        End Property


#End Region

#Region "Constructors"


        'Empty Consructor
        Public Sub New()

        End Sub

        'Use this Constructor when wanting to set the ProductCategoryID on Creation
        Public Sub New(ByVal ProductCategoryID As Integer)

            _lngProductCategoryID = ProductCategoryID 'Sets Local ProductID Variable

        End Sub




        ' Use this Constructor when wanting to create a NEW Productcategory(Requires a "sProductcategory") 
        Public Sub New(ByRef sProductCategory As sProductCategory, ByRef CreatedProductCategory As Boolean)

            CreatedProductCategory = AddProductCategory(sProductCategory)

        End Sub




#End Region

#Region "Functions & Procedures"


        Public Function RefreshShippingRatesForCBSOA(ShippingPolicyXML As XmlDocument) As Boolean

            Dim bSuccess As Boolean 'Return variable 
            Dim Tier As XmlElement
            Dim Charge As XmlElement
            Dim Policy As XmlElement



            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseCBSOADBConnection").ToString)

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
                    SQLCommand.CommandText = "CBSOA..[TEMP_DeleteShippingRatesByPolicyID]"     'Stored Procedure Name


                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShippingPolicyID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CInt(Policy.Attributes("policyID").Value)))

                    SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                    'bSuccess = True

                    'Reset CMD for next SP
                    SQLCommand.CommandText = "CBSOA..[TEMP_AddShippingCharge]"     'Stored Procedure Name


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
                                    _strErrorMessage = ex.ToString       'Set the Classes Error Message
                                End Try



                            Next
                        End If

                    Next

                Next

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
                                    _strErrorMessage = ex.ToString       'Set the Classes Error Message
                                End Try



                            Next
                        End If

                    Next

                Next

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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



        Public Function AddLLRepPortalCategory(ByVal sProductCategory As sProductCategory) As Boolean

        End Function


        Public Function AddProductCategory(ByVal sProductCategory As sProductCategory) As Boolean

            'Function: AddProductCategory ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductCategory
            '       Parent Class:  ProductCategory
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductCategory Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover 
            '       Created on:  11/16/05
            '       Description: Adds a Product Category
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_AddProductCategory]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure

                With sProductCategory

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Name))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .BrandID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 1000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Description))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .PrimaryResourceImageID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CreatedBy))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Active))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ExternalCategoryID", System.Data.SqlDbType.VarChar, 1000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ExternalCategoryID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ExternalCategorySource", System.Data.SqlDbType.VarChar, 1000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ExternalCategorySource))

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                End With
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                If Not SQLCommand.Parameters("@ProductCategoryID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product Category(Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else

                    sProductCategory.ProductCategoryID = SQLCommand.Parameters("@ProductCategoryID").Value

                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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



        Public Function UpdateProductCategory(ByVal sProductCategory As sProductCategory) As Boolean

            'Function: UpdateProductCategory ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateProductCategory
            '       Parent Class:  ProductCategory
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductCategory Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover 
            '       Created on:  11/16/05
            '       Description: Updates a Product Category
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_UpdateProductCategory]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure

                With sProductCategory

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductCategoryID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Name))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 1000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Description))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .Active))

                End With

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure
                bSuccess = True


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function AddFeaturedProduct(ByVal ProductSeriesID As Integer, ByVal ProductCategoryID As Integer) As Boolean

            'Function: AddFeaturedProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddFeaturedProduct
            '       Parent Class:  ProductCategory
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductCategory Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Corey Perrymond
            '       Created on:  4/23/07
            '       Description: Adds a Product as a Featured Product to a Category
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 
            Dim ReturnValue As Integer ' SQL return value

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "[AddFeaturedProduct]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure



                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductSeriesID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductCategoryID))



                ReturnValue = SQLCommand.ExecuteScalar()    'Execute the Stored Procedure
                If ReturnValue > 0 Then
                    bSuccess = True
                ElseIf ReturnValue = 0 Then
                    bSuccess = False
                    Throw New Exception("Internal Error adding Featured Product(Not System Error)")
                ElseIf ReturnValue = -1 Then
                    Throw New Exception("Featured Product already exists for the specified Category")      'Set the Classes Error Message
                    bSuccess = False
                End If


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function AddProductCategoryBrandJunc(ByVal sProductCategory As sProductCategory) As Boolean

            'Function: AddProductCategoryBrandJunc ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductCategoryBrandJunc
            '       Parent Class:  ProductCategory
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductcategory Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover 
            '       Created on:  11/16/05
            '       Description: Adds a Product Category Brand JUNC
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_AddProductCategoryBrandJUNC]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure

                With sProductCategory

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductCategoryID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ParentCategoryID", System.Data.SqlDbType.Int, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ParentCategoryID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryJUNCID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                End With
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                If Not SQLCommand.Parameters("@ProductCategoryJUNCID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product Category JUNC (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else

                    sProductCategory.ProductCategoryBrandID = SQLCommand.Parameters("@ProductCategoryJUNCID").Value

                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function DeleteAllFeaturedProducts(ByVal ProductCategoryID As Integer) As Boolean

            'Function: DeleteAllFeaturedProducts ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteAllFeaturedProducts
            '       Parent Class:  ProductCategory
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductCategory Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Corey Perrymond
            '       Created on:  4/23/07
            '       Description: Deletes all Featured Products for a Category
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 
            Dim ReturnValue As Integer ' SQL return value

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "[DeleteFeaturedProductsByCategory]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductCategoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductCategoryID))



                ReturnValue = SQLCommand.ExecuteScalar()    'Execute the Stored Procedure
                If ReturnValue > 0 Then
                    bSuccess = True
                ElseIf ReturnValue = 0 Then
                    bSuccess = False
                    Throw New Exception("Internal Error deleting Featured Products(Not System Error)")
                ElseIf ReturnValue = -1 Then
                    Throw New Exception("Featured Product does not exists for the specified Category")      'Set the Classes Error Message
                    bSuccess = False
                End If


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function DeleteFeaturedProduct(ByVal ProductSeriesID As Integer) As Boolean

            'Function: DeleteFeaturedProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteFeaturedProduct
            '       Parent Class:  ProductCategory
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductCategory Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Corey Perrymond
            '       Created on:  4/23/07
            '       Description: Deletes a Product Series from Feature Product Table 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 
            Dim ReturnValue As Integer ' SQL return value

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product Series One Step is Required: ********************
            '   - 1. Add the Product Series table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "[DeleteFeaturedProduct]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductSeries" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ProductSeriesID))



                ReturnValue = SQLCommand.ExecuteScalar()    'Execute the Stored Procedure
                If ReturnValue > 0 Then
                    bSuccess = True
                ElseIf ReturnValue = 0 Then
                    bSuccess = False
                    Throw New Exception("Internal Error deleting Featured Products(Not System Error)")
                ElseIf ReturnValue = -1 Then
                    Throw New Exception("Featured Product does not exists for the specified Category")      'Set the Classes Error Message
                    bSuccess = False
                End If


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function ResetAllEcomCategoryies(ByVal BrandID As Integer) As Boolean

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To delete Product Category JUNC One Step is Required: ********************
            '   - 1. Delete ALL ProductCategory JUNC
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[r_ResetAllEcomCategoryies]"     'Stored Procedure Name

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, BrandID))


                SQLCommand.ExecuteNonQuery()
                bSuccess = True


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

 

        Public Function GetProductCategory(ByRef sProductCategory As sProductCategory) As Boolean

            'Function: GetProductCategory ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetProductCategory
            '       Parent Class:  ProductCategory
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductCategory Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  04/17/07
            '       Description: Gets a Product Category
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            Dim blnIncludeJuncInfo As Boolean = False
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database

                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                SQLCommand.Connection = SQLConn
                'Set the Connection
                SQLCommand.CommandText = "LLFBPS..[r_GetProductCategory]"              'Stored Procedure Name


                If sProductCategory.ExternalCategoryID <> "" AndAlso sProductCategory.BrandID > 0 Then

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ExternalCategoryID", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProductCategory.ExternalCategoryID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProductCategory.BrandID))

                End If


                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Category Structure
                SQLDataReader.Read()
                If SQLDataReader.HasRows Then

                    With sProductCategory
                        'Fill  Information
                        If Not IsDBNull(SQLDataReader("ProductCategoryID")) Then .ProductCategoryID = SQLDataReader("ProductCategoryID")
                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Product Category Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If
                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            Return bSuccess

        End Function





#End Region
    End Class


    Public Class Image

        '"Image" Class Profile ******************************************************
        '
        '       Class: Image 
        '       Parent File: WorkElements.vb
        '       Parent Project: CMS
        '       Parent Solution: CMS
        '       Created By: Amy Cook
        '       Created on: 01/04/06
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs
#End Region

#Region "Class Properties"


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage()
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber()
            Get
                Return _intErrorNumber
            End Get
        End Property


#End Region

#Region "Constructors"


        'Empty Consructor
        Public Sub New()

        End Sub


#End Region

#Region "Functions & Procedures"
        Public Function GetImage(ByRef sImage As sImage) As Boolean

            'Function: GetImage ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetImage
            '       Parent Class:  Image
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sImage Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/09/06
            '       Description: Gets an Image
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


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

                SQLCommand.CommandText = "[GetImages]"

                'Set the ProdcutID Parameter
                If Not sImage.ResourceTypeID > 0 And Not sImage.ResourceID > 0 Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sImage.ImageID))
                End If

                If sImage.ResourceTypeID > 0 Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sImage.ResourceTypeID))
                End If

                If sImage.ResourceID > 0 Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sImage.ResourceID))
                End If



                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Image Structure

                If SQLDataReader.Read Then

                    With sImage
                        'Fill Image Information
                        If Not IsDBNull(SQLDataReader("ImageID")) Then .ImageID = SQLDataReader("ImageID")
                        If Not IsDBNull(SQLDataReader("Name")) Then .Description = SQLDataReader("Name")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        If Not IsDBNull(SQLDataReader("Filename")) Then .Filename = SQLDataReader("Filename")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Image Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If
                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            Return bSuccess

        End Function
        Public Function UpdateImage(ByRef sImage As sImage) As Boolean

            'Function: UpdateImage ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateImage
            '       Parent Class:  Image
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sImage Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/09/06
            '       Description: Updates an Image
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


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

                SQLCommand.CommandText = "BPSv2..[UpdateImageDescription]"

                'Set the ProdcutID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sImage.ImageID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sImage.Description))


                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Updating Image (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                End If

                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()



            Catch SQLErr As SqlException
                bSuccess = False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                bSuccess = False
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            Return bSuccess

        End Function

        Public Function AddImage(ByRef sImage As sImage) As Boolean

            'Function: AddImage ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddImage
            '       Parent Class:  Image
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): AddImage Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/04/06
            '       Description: Adds a Image
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Image - One Step is Required: ********************
            '   - 1. Add the Image table entry and ImageResourceJUNC entry (all in one sp call)
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "[AddImage]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sImage" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImagePath", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@Name").Value = sImage.Name
                SQLCommand.Parameters("@Description").Value = sImage.Description
                SQLCommand.Parameters("@ImagePath").Value = sImage.ImagePath
                SQLCommand.Parameters("@ResourceTypeID").Value = sImage.ResourceTypeID
                SQLCommand.Parameters("@ResourceID").Value = sImage.ResourceID
                SQLCommand.Parameters("@CompanyID").Value = sImage.CompanyID
                SQLCommand.Parameters("@BrandID").Value = sImage.BrandID
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "ImageID" back to the sImage Structure
                sImage.ImageID = SQLCommand.Parameters("@ImageID").Value

                If Not SQLCommand.Parameters("@ImageID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Image (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function AddAssociatedImage(ByRef sImage As sImage) As Boolean

            'Function: AddPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddImage
            '       Parent Class:  Image
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): AddImage Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/04/06
            '       Description: Adds a Image
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Image - One Step is Required: ********************
            '   - 1. Add the Image table entry and ImageResourceJUNC entry (all in one sp call)
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "[AddAssociatedImage]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sImage" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImagePath", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@Name").Value = sImage.Name
                SQLCommand.Parameters("@Description").Value = sImage.Description
                SQLCommand.Parameters("@ImagePath").Value = sImage.ImagePath
                SQLCommand.Parameters("@ResourceTypeID").Value = sImage.ResourceTypeID
                SQLCommand.Parameters("@ResourceID").Value = sImage.ResourceID
                SQLCommand.Parameters("@CompanyID").Value = sImage.CompanyID
                SQLCommand.Parameters("@BrandID").Value = sImage.BrandID
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "ImageID" back to the sImage Structure
                sImage.ImageID = SQLCommand.Parameters("@ImageID").Value

                If Not SQLCommand.Parameters("@ImageID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Image (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                Else
                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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


        Public Function AddImageToResource(ByVal CompanyID As Integer, ByVal BrandID As Integer, ByVal ResourceTypeID As Integer, ByVal ResourceID As Integer, ByVal ImageId As Integer) As Boolean

            'Function: AddPart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddImageToResource
            '       Parent Class:  Image
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ResourceTypeID
            '       Required Parameters(1): ResourceID
            '       Required Parameters(1): ImageID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/04/06
            '       Description: Adds a Image to a Resource
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Image - One Step is Required: ********************
            '   - 1. Add the ImageResourceJUNC entry (or if Primary, add directly to Resource Table
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "[AddImageToResource]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sImage" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))



                SQLCommand.Parameters("@ResourceTypeID").Value = ResourceTypeID
                SQLCommand.Parameters("@ResourceID").Value = ResourceID
                SQLCommand.Parameters("@ImageID").Value = ImageId
                SQLCommand.Parameters("@CompanyID").Value = CompanyID
                SQLCommand.Parameters("@BrandID").Value = BrandID
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Adding Image Resource (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function SetNewPrimaryResourceImage(ByVal ResourceTypeID As Integer, ByVal ResourceID As Integer, ByVal NewPrimaryResourceImageID As Integer) As Boolean

            'Function: SetNewPrimaryResourceImage ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: SetNewPrimaryResourceImage
            '       Parent Class:  Image
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ResourceTypeID 
            '       Required Parameters(1): ResourceID
            '       Required Parameters(1): NewPrimaryResourceImageID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/05/06
            '       Description: Reset's the PrimaryResourceImage
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[SetNewPrimaryResourceID]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values 

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@NewPrimaryResourceImageID", System.Data.SqlDbType.Int, 4))


                SQLCommand.Parameters("@ResourceTypeID").Value = ResourceTypeID
                SQLCommand.Parameters("@ResourceID").Value = ResourceID
                SQLCommand.Parameters("@NewPrimaryResourceImageID").Value = NewPrimaryResourceImageID
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Setting PrimaryResourceImageID (Not System Error)"     'Set the Classes Error Message
                    _strErrorMessage = 0      'Set the Classes Error Number
                    bSuccess = False

                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False


                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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
        Public Function DeleteImageResource(ByVal ImageResourceJunctionID As Long) As Boolean

            'Function: DeleteImageResource ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteImageResource
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ImageResourceJunctionID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/17/05
            '       Description: Deletes a ImageResourceJunc entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "BPSv2..[DeleteImageResourceJUNC]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageResourceJunctionID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@ImageResourceJunctionID").Value = ImageResourceJunctionID

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                Return True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

        Public Function DeleteImage(ByVal ImageID As Long) As Boolean

            'Function: DeleteImageResource ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteImageResource
            '       Parent Class:  Part
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ImageID
            '       Optional Parameters(0):
            '
            '       Created By:  Corey perrymond
            '       Created on:  4/11/07
            '       Description: Deletes a ImageID and related Image Associations
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "[DeleteImage]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@ImageID").Value = ImageID

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                Return True

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                _strErrorMessage = SQLErr.ToString       'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number       'Set the Classes Error Number
                bSuccess = False

                'MISC ERROR CATCH
            Catch Err As Exception
                _strErrorMessage = Err.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
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

#End Region
    End Class







    Public Structure sProduct
        Public ProductID As Long
        Public ProductNumber As String
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
        Public Year As Integer
        Public MSRP As Double
        Public IsSellable As Boolean
        Public Weight As String
        Public UPCCode As String
        Public AvailableQuantity As Integer
        Public ExpectedDate As Date
        Public CreatedBy As Integer
        Public Surcharge As Double
        Public CreatedByName As String
        Public ProductJunctionID As Integer
        Public ProductSeriesID As Integer
        Public ProductSeriesName As String
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public DateCreated As Date
        Public IsActive As Boolean
        Public UserField1 As String
        Public UserField2 As String
        Public UserField3 As String
        Public UserField4 As String
        Public UserField5 As String
        Public _LogInformation
        Public _UserInformation
    End Structure
    Public Structure sProductSeries
        Public ProductSeriesID As Integer
        Public Name As String
        Public Description As String
        Public WebDescription As String
        Public Freight As String
        Public IsSupported As Boolean
        Public Specifications As String
        Public WarrantyInfo As String
        Public Bonus As Integer
        Public ImagePath As String
        Public PrimaryResourceImageID As Long
        Public PrimaryResourceImageIDLarge As Long
        Public PrimaryResourceImageIDSmall As Long
        Public PrimaryResourceImageIDFeature As Long
        Public PrimaryResourceImageName As String
        Public PrimaryResourceImageNameLarge As String
        Public PrimaryResourceImageNameSmall As String
        Public PrimaryResourceImageNameFeature As String
        Public ImagePathLarge As String
        Public ImagePathSmall As String
        Public ImagePathFeature As String

        Public AvailableQuantity As Integer
        Public ExpectedDate As Date

        Public CreatedBy As Integer
        Public CreatedByName As String
        Public DateCreated As Date
        Public Active As Boolean

        Public ProductSeriesNumber As String
        Public IsBundle As Boolean
        Public MSRP As String 'Changed to string to hold price ranges
        Public AdditionalShipping As Double
        Public Multiplier As Integer
        Public HasSubProducts As Boolean
        Public IsOversized As Boolean

        Public IsPersonalized As Boolean
        Public PersonalizationDescription As String

        Public UserField1 As String
        Public UserField2 As String

        Public ProductSeriesJunctionID As Integer
        Public CompanyID As Integer
        Public CompanyName As String
        Public BrandID As Integer
        Public BrandName As String
        Public ProductCategoryID As Integer
        Public ProductCategoryName As String
        Public ProductID As Long
        Public _LogInformation
        Public _UserInformation

        Public ProductGroupName As String
        Public AttributeTypeID As Integer
        Public AttributeTypeName As String
        Public IsSellable As Boolean
        Public CaseQty As Integer
        Public IsCaseQTYOnly As Boolean
        Public IsHidden As Boolean
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

        Public _LogInformation
        Public _UserInformation
    End Structure


    Public Structure sProductCategory
        Public ProductCategoryID As Integer
        Public Name As String
        Public Description As String
        Public ParentCategoryID As Integer
        Public ParentCategoryName As String
        Public PrimaryResourceImageID As Long
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
        Public ExternalCategoryID As String
        Public ExternalCategorySource As String
        Public ProductCategoryBrandID As Long
        Public _LogInformation
        Public _UserInformation

    End Structure

    Public Class SectionCategory
        Public SectionCategoryID As Integer
        Public SectionID As Integer
        Public CategoryID As Integer?
        Public Description As String
        Public ParentSectionID As Integer?
    End Class

    Public Structure sProductSectionCategory
        Public ProductNumber As String
        Public SectionID As Integer
    End Structure

    Public Structure sImage
        Public ImageID As Integer
        Public Name As String
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

    Public Structure sProductSeriesAttribute
        Public ProductSeriesID As Integer
        Dim AttributeTypeID As Integer
        Dim AttributeName As String
        Dim AttributeValue As String
        Dim DateCreated As Date
        Dim DateModified As Date


    End Structure





    Public Enum ImageType
        Alternate = 1
        DataCube = 2

    End Enum
    Public Enum ProductSeriesAssociatedType
        CrossSell = 3
        'UpSell = 2
        Bundled = 5

    End Enum

    Public Enum ProductSeriesAttributeType
        Bullet = 1
        DropShip = 2
        TruckDelivery = 3
        CompanyExclusive = 4

        Oversized = 5
        Length = 6
        Width = 7
        Height = 8
        QtyPer = 9
        Scent = 10
        Weight = 11
        UsageLocation = 19
        Size = 20
        Color = 21
        InStoreOnly = 22
        NewProduct = 23

    End Enum
End Namespace