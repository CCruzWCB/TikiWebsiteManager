Imports System.Net.Mail
Imports System.Text
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Configuration
Imports System.Web.Security
Imports System.Data


Namespace BPS_BL.BPS

    Public Class Company

        '"Company" Class Profile ******************************************************
        '
        '       Class: Company 
        '       Parent File: WorkElements.vb
        '       Parent Project: CMS
        '       Parent Solution: CMS
        '       Created By: Amy Cook
        '       Created on: 07/21/05
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _lngCompamyID As Long          'Used to Hold the CompanyID currently being Used/Created
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs
#End Region

#Region "Class Properties"


        'Gets & Sets Current CompamyID
        Public Property CompanyID() As Integer
            Get
                Return _lngCompamyID  'Gets Local _lngProductID Variable
            End Get
            Set(ByVal Value As Integer)
                _lngCompamyID = Value 'Sets Local _lngProductID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
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
        Public Sub New(ByVal CompanyID As Integer)

            _lngCompamyID = CompanyID 'Sets Local ProductID Variable

        End Sub




        ' Use this Constructor when wanting to create a NEW Company (Requires a "sCompany") 
        Public Sub New(ByRef sCompany As sCompany, ByRef CreatedCompany As Boolean)

            'CreatedCompany = AddCompany(sCompany)

        End Sub




#End Region

#Region "Functions & Procedures"
        'Public Function AddCompany()

        'End Function


#End Region
    End Class

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
        Public Property ProductID() As Long
            Get
                Return _lngProductID  'Gets Local _lngProductID Variable
            End Get
            Set(ByVal Value As Long)
                _lngProductID = Value 'Sets Local _lngProductID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
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




        ' Use this Constructor when wanting to create a NEW Product(Requires a "sProduct") 
        Public Sub New(ByRef sProduct As sProduct, ByRef CreatedProduct As Boolean)

            CreatedProduct = AddProduct(sProduct)

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
                SQLCommand.CommandText = "LLFBPS..[AddProduct]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductModelNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Year", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Freight", System.Data.SqlDbType.VarChar, 500, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSupported", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Specifications", System.Data.SqlDbType.VarChar, 1250, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSellable", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FirstCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SecondCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@ProductModelNumber").Value = sProduct.ProductModelNumber
                SQLCommand.Parameters("@Name").Value = sProduct.Name
                SQLCommand.Parameters("@Description").Value = sProduct.Description
                SQLCommand.Parameters("@Year").Value = sProduct.ProductYear
                SQLCommand.Parameters("@Freight").Value = sProduct.Freight
                SQLCommand.Parameters("@IsSupported").Value = sProduct.IsSupported
                SQLCommand.Parameters("@Specifications").Value = sProduct.Specifications

                If sProduct.MSRP > 0 Then
                    SQLCommand.Parameters("@MSRP").Value = sProduct.MSRP
                Else
                    SQLCommand.Parameters("@MSRP").Value = DBNull.Value
                End If

                SQLCommand.Parameters("@IsSellable").Value = sProduct.IsSellable

                SQLCommand.Parameters("@PrimaryResourceImageID").Value = sProduct.PrimaryResourceImageID
                SQLCommand.Parameters("@CreatedBy").Value = sProduct.CreatedBy
                SQLCommand.Parameters("@CreatedName").Value = sProduct.CreatedByName
                SQLCommand.Parameters("@Active").Value = sProduct.IsActive

                SQLCommand.Parameters("@CompanyID").Value = sProduct.CompanyID
                SQLCommand.Parameters("@BrandID").Value = sProduct.BrandID
                SQLCommand.Parameters("@FirstCategoryID").Value = sProduct.ProductCategory1
                SQLCommand.Parameters("@SecondCategoryID").Value = sProduct.ProductCategory2
                SQLCommand.Parameters("@ProductSeriesID").Value = sProduct.ProductSeriesID
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

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
        Public Function GetProductDiagram(ByRef sProductDiagram As sProductDiagram) As Boolean

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

                SQLCommand.CommandText = "LLFBPS..[GetProductDiagram]"


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProductDiagram.ProductID))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)

                'Populate the Image Structure

                If SQLDataReader.Read Then

                    With sProductDiagram
                        'Fill Image Information
                        If Not IsDBNull(SQLDataReader("ProductDiagramID")) Then .ProductDiagramID = SQLDataReader("ProductDiagramID")
                        If Not IsDBNull(SQLDataReader("ProductID")) Then .ProductID = SQLDataReader("ProductID")
                        If Not IsDBNull(SQLDataReader("Title")) Then .Title = SQLDataReader("Title")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("FilePath")) Then .FilePath = SQLDataReader("FilePath")
                        If Not IsDBNull(SQLDataReader("FileName")) Then .FileName = SQLDataReader("FileName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedByName")) Then .CreatedByName = SQLDataReader("CreatedByName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                    End With
                Else
                    bSuccess = False

                End If
                'Cleanup
                SQLDataReader.Close()
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
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            Return bSuccess

        End Function
        Public Function DeleteAllProductDiagrams(ByVal ProductID As Integer) As Boolean

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

                SQLCommand.CommandText = "LLFBPS..[DeleteALLProductDiagrams]"


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ProductID))

                'Set the DataReader
                SQLCommand.ExecuteReader(CommandBehavior.SingleResult)



                'Cleanup
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
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            Return bSuccess

        End Function

        Public Function AddProductDiagram(ByRef sProductDiagram As sProductDiagram) As Boolean

            'Function: AddProductDiagram ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductDiagram
            '       Parent Class:  Product
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sProductDiagram Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  09/18/06
            '       Description: Adds a Product Diagram
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add ProductDiagram One Step is Required: ********************
            '   - 1. Add the ProductDiagram table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[AddProductDiagram]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sProductDiagram" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Title", System.Data.SqlDbType.VarChar, 75))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FilePath", System.Data.SqlDbType.VarChar, 500))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FileName", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductDiagramID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@ProductID").Value = sProductDiagram.ProductID
                SQLCommand.Parameters("@Title").Value = sProductDiagram.Title
                SQLCommand.Parameters("@Description").Value = sProductDiagram.Description
                SQLCommand.Parameters("@FilePath").Value = sProductDiagram.FilePath
                SQLCommand.Parameters("@FileName").Value = sProductDiagram.FileName
                SQLCommand.Parameters("@DateCreated").Value = sProductDiagram.DateCreated
                SQLCommand.Parameters("@CreatedBy").Value = sProductDiagram.CreatedBy


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "ProductDiagramID" back to the sProductDiagram Structure
                sProductDiagram.ProductDiagramID = SQLCommand.Parameters("@ProductDiagramID").Value

                If Not SQLCommand.Parameters("@ProductDiagramID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product Diagram (Not System Error)"     'Set the Classes Error Message
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

        Public Function AddProduct_InheritFromSeries(ByRef sProduct As sProduct) As Boolean

            'Function: AddProduct_InheritFromSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProduct_InheritFromSeries
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
            '       Created on:  11/30/05
            '       Description: Adds a Product (Inheriting from Series)
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


                'Verify that ProductSeriesID exists
                If sProduct.ProductSeriesID > 0 Then

                    SQLConn.Open()    'Open Database


                    'Set the Basic Command Information 
                    SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                    SQLCommand.Connection = SQLConn       'Set the Connection


                    'Set the Specific Command Information 
                    SQLCommand.CommandText = "LLFBPS..[AddProduct_InheritFromSeries]"     'Stored Procedure Name

                    'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure


                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductModelNumber", System.Data.SqlDbType.VarChar, 50))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedByName", System.Data.SqlDbType.VarChar, 100))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                    SQLCommand.Parameters("@ProductSeriesID").Value = sProduct.ProductSeriesID
                    SQLCommand.Parameters("@ProductModelNumber").Value = sProduct.ProductModelNumber
                    SQLCommand.Parameters("@Name").Value = sProduct.Name
                    SQLCommand.Parameters("@CreatedBy").Value = sProduct.CreatedBy
                    SQLCommand.Parameters("@CreatedByName").Value = sProduct.CreatedByName
                    SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                    'Pass the new "ProductID" back to the sProduct Structure
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
                Else
                    _strErrorMessage = "Unable to Add Model.  ProductSeriesID is a required field." 'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number
                End If

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

        Public Function UpdateProduct(ByVal sProduct As sProduct) As Boolean

            'Function: UpdateProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateProduct
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
            '       Description: Updates a Product 
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
                SQLCommand.CommandText = "LLFBPS..[UpdateProduct]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductModelNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 1000))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Year", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Freight", System.Data.SqlDbType.VarChar, 500))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSupported", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4))


                SQLCommand.Parameters("@ProductID").Value = sProduct.ProductID
                SQLCommand.Parameters("@ProductModelNumber").Value = sProduct.ProductModelNumber
                SQLCommand.Parameters("@Name").Value = sProduct.Name
                SQLCommand.Parameters("@Description").Value = sProduct.Description
                SQLCommand.Parameters("@Year").Value = sProduct.ProductYear
                SQLCommand.Parameters("@Freight").Value = sProduct.Freight
                SQLCommand.Parameters("@IsSupported").Value = sProduct.IsSupported
                SQLCommand.Parameters("@Specifications").Value = sProduct.Specifications
                SQLCommand.Parameters("@Active").Value = sProduct.IsActive

                'SQLCommand.Parameters("@ProductJunctionID").Value = sProduct.ProductJunctionID
                'SQLCommand.Parameters("@CompanyID").Value = sProduct.CompanyID
                'SQLCommand.Parameters("@BrandID").Value = sProduct.BrandID
                'SQLCommand.Parameters("@FirstCategoryID").Value = sProduct.ProductCategory1
                'SQLCommand.Parameters("@SecondCategoryID").Value = sProduct.ProductCategory2
                'SQLCommand.Parameters("@ProductSeriesID").Value = sProduct.ProductSeriesID
                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Updating Product (Not System Error)"     'Set the Classes Error Message
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
        Public Function UpdateProductImages(ByVal sProduct As sProduct) As Boolean

            'Function: UpdateProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateProductImages
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
            '       Created on:  08/27/07
            '       Description: Updates a Product's images
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
                SQLCommand.CommandText = "LLFBPS..[UpdateProductImages]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_large", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_small", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_feature", System.Data.SqlDbType.Int, 4))


                SQLCommand.Parameters("@ProductID").Value = sProduct.ProductID
                SQLCommand.Parameters("@PrimaryResourceImageID").Value = sProduct.PrimaryResourceImageID
                SQLCommand.Parameters("@PrimaryResourceImageID_large").Value = sProduct.PrimaryResourceImageIDLarge
                SQLCommand.Parameters("@PrimaryResourceImageID_small").Value = sProduct.PrimaryResourceImageIDSmall
                SQLCommand.Parameters("@PrimaryResourceImageID_feature").Value = sProduct.PrimaryResourceImageIDFeature


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Updating Product Images (Not System Error)"     'Set the Classes Error Message
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


        Public Function MigrateProductToSeries(ByVal ProductID As Integer, ByVal NewProductSeriesID As Integer, ByVal CurrentProductSeriesID As Integer, ByVal CreatedBy As Integer, ByVal CreatedByName As String) As Boolean

            'Function: MigrateProductToSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: MigrateProductToSeries
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
            '       Description: Migrates a Product To another Series
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
                SQLCommand.CommandText = "LLFBPS..[MigrateProductToAnotherSeries]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@NewProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CurrentProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedByName", System.Data.SqlDbType.VarChar, 100))


                SQLCommand.Parameters("@ProductID").Value = ProductID
                SQLCommand.Parameters("@NewProductSeriesID").Value = NewProductSeriesID
                SQLCommand.Parameters("@CurrentProductSeriesID").Value = CurrentProductSeriesID
                SQLCommand.Parameters("@CreatedBy").Value = CreatedBy
                SQLCommand.Parameters("@CreatedByName").Value = CreatedByName

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Migrating Product to Series (Not System Error)"     'Set the Classes Error Message
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

            If sProduct.ProductID > 0 Then
                _lngProductID = sProduct.ProductID
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

                SQLCommand.CommandText = "[GetProduct]"

                'Set the ProdcutID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PriceListID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProduct.PriceListID))

                ''Check to see if the BrandID was passed
                ''If sProduct.BrandID > 0 Then 'GET BY BRAND & Product
                ''    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProduct.BrandID))
                ''End If

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sProduct
                        'Fill Knowledge Information
                        If Not IsDBNull(SQLDataReader("ProductID")) Then .ProductID = SQLDataReader("ProductID")
                        If Not IsDBNull(SQLDataReader("ProductModelNumber")) Then .ProductModelNumber = SQLDataReader("ProductModelNumber")
                        If Not IsDBNull(SQLDataReader("ProductSeriesID")) Then .ProductSeriesID = SQLDataReader("ProductSeriesID")
                        If Not IsDBNull(SQLDataReader("SeriesName")) Then .ProductSeriesName = SQLDataReader("SeriesName")
                        If Not IsDBNull(SQLDataReader("Year")) Then .ProductYear = SQLDataReader("Year")
                        If Not IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("Freight")) Then .Freight = SQLDataReader("Freight")
                        If Not IsDBNull(SQLDataReader("IsSupported")) Then .IsSupported = SQLDataReader("IsSupported")
                        If Not IsDBNull(SQLDataReader("Specifications")) Then .Specifications = SQLDataReader("Specifications")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID_Large")) Then .PrimaryResourceImageIDLarge = SQLDataReader("PrimaryResourceImageID_Large")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID_Small")) Then .PrimaryResourceImageIDSmall = SQLDataReader("PrimaryResourceImageID_Small")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID_Feature")) Then .PrimaryResourceImageIDFeature = SQLDataReader("PrimaryResourceImageID_Feature")
                        If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        If Not IsDBNull(SQLDataReader("ImagePath_large")) Then .ImagePathLarge = SQLDataReader("ImagePath_large")
                        If Not IsDBNull(SQLDataReader("ImagePath_small")) Then .ImagePathSmall = SQLDataReader("ImagePath_small")
                        If Not IsDBNull(SQLDataReader("ImagePath_feature")) Then .ImagePathFeature = SQLDataReader("ImagePath_feature")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("Active")) Then .IsActive = SQLDataReader("Active")
                        If Not IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
                        If Not IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
                        If Not IsDBNull(SQLDataReader("ParentCategoryID")) Then .ProductCategory1 = SQLDataReader("ParentCategoryID")
                        If Not IsDBNull(SQLDataReader("ParentCategoryName")) Then .ProductCategory1Name = SQLDataReader("ParentCategoryName")
                        If Not IsDBNull(SQLDataReader("ProductCategoryID")) Then .ProductCategory2 = SQLDataReader("ProductCategoryID")
                        If Not IsDBNull(SQLDataReader("CategoryName")) Then .ProductCategory2Name = SQLDataReader("CategoryName")
                        If Not IsDBNull(SQLDataReader("MSRP")) Then .MSRP = SQLDataReader("MSRP")
                        If Not IsDBNull(SQLDataReader("IsSellable")) Then .IsSellable = SQLDataReader("IsSellable")
                        If Not IsDBNull(SQLDataReader("AvailableQuantity")) Then .AvailableQuantity = SQLDataReader("AvailableQuantity")
                        If Not IsDBNull(SQLDataReader("ExpectedDate")) Then .ExpectedDate = SQLDataReader("ExpectedDate")
                        If Not IsDBNull(SQLDataReader("AdditionalShipping")) Then .AdditionalShipping = SQLDataReader("AdditionalShipping")


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


        Public Function GetProduct(ByRef sProduct As sProduct, ByVal ProductModelNumber As String, Optional ByVal CompanyID As Long = 0, Optional ByVal BrandID As Long = 0) As Boolean

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

                SQLCommand.CommandText = "LLFBPS..[GetProductByModelNumber]"

                'Set the Product Model Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductModelNumber", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ProductModelNumber))



                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)

                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sProduct
                        'Fill Knowledge Information
                        If Not IsDBNull(SQLDataReader("ProductID")) Then .ProductID = SQLDataReader("ProductID")
                        If Not IsDBNull(SQLDataReader("ProductModelNumber")) Then .ProductModelNumber = SQLDataReader("ProductModelNumber")
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
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("Active")) Then .IsActive = SQLDataReader("Active")
                        If Not IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
                        If Not IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
                        If Not IsDBNull(SQLDataReader("ParentCategoryID")) Then .ProductCategory1 = SQLDataReader("ParentCategoryID")
                        If Not IsDBNull(SQLDataReader("ParentCategoryName")) Then .ProductCategory1Name = SQLDataReader("ParentCategoryName")
                        If Not IsDBNull(SQLDataReader("ProductCategoryID")) Then .ProductCategory2 = SQLDataReader("ProductCategoryID")
                        If Not IsDBNull(SQLDataReader("CategoryName")) Then .ProductCategory2Name = SQLDataReader("CategoryName")
                        If Not IsDBNull(SQLDataReader("MSRP")) Then .MSRP = SQLDataReader("MSRP")
                        If Not IsDBNull(SQLDataReader("IsSellable")) Then .IsSellable = SQLDataReader("IsSellable")

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



        Public Function AvailableSerialNumber(ByRef sRegisteredSerialNumber As sRegisteredSerialNumber) As Boolean

            'Function: AvailableSerialNumber ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AvailableSerialNumber
            '       Parent Class:  Product
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sRegisteredSerialNumber Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  03/27/06
            '       Description: Checks to see if a Serial/Product Number have bee Registered
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = False  'Used to confirm that the Function Succesfully Processed

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database

                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                          'Set the Connection

                SQLCommand.CommandText = "LLFBPS..[GetRegisteredSerialNumber]"

                'Set the ProdcutID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sRegisteredSerialNumber.CompanyID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sRegisteredSerialNumber.BrandID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sRegisteredSerialNumber.ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SerialNumber", System.Data.SqlDbType.VarChar, 150, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sRegisteredSerialNumber.SerialNumber))



                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then
                    bSuccess = False

                    With sRegisteredSerialNumber
                        'Fill Information
                        If Not IsDBNull(SQLDataReader("ProductID")) Then .ProductID = SQLDataReader("ProductID")
                        If Not IsDBNull(SQLDataReader("ProductModelNumber")) Then .ProductModelNumber = SQLDataReader("ProductModelNumber")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("CompanyID")) Then .CompanyID = SQLDataReader("CompanyID")
                        If Not IsDBNull(SQLDataReader("CompanyName")) Then .CompanyName = SQLDataReader("CompanyName")
                        If Not IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
                        If Not IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
                        If Not IsDBNull(SQLDataReader("CustomerID")) Then .CustomerID = SQLDataReader("CustomerID")
                        If Not IsDBNull(SQLDataReader("CustomerName")) Then .CustomerName = SQLDataReader("CustomerName")
                        If Not IsDBNull(SQLDataReader("ProductName")) Then .ProductName = SQLDataReader("ProductName")
                        If Not IsDBNull(SQLDataReader("CustomerOwnershipID")) Then .CustomerOwnershipID = SQLDataReader("CustomerOwnershipID")
                        If Not IsDBNull(SQLDataReader("RegisteredSerialNumberID")) Then .RegisteredSerialNumberID = SQLDataReader("RegisteredSerialNumberID")
                        If Not IsDBNull(SQLDataReader("SerialNumber")) Then .SerialNumber = SQLDataReader("SerialNumber")
                        If Not IsDBNull(SQLDataReader("SerialNumberConfirmed")) Then .SerialNumberConfirmed = SQLDataReader("SerialNumberConfirmed")

                    End With
                Else
                    bSuccess = True
                    _strErrorMessage = "Serial Number Not Found"                   'Set the Classes Error Message
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



        Public Function AddRegisteredSerialNumber(ByRef sRegisteredSerialNumber As sRegisteredSerialNumber) As Boolean

            'Function: AddRegisteredSerialNumber ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddRegisteredSerialNumber
            '       Parent Class:  Product
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sRegisteredSerialNumber Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  03/16/05
            '       Description: Adds a Registered Serial Number
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add  One Step is Required: ********************
            '   - 1. Add the Registered Serial Number
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[AddRegisteredSerialNumber]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                With sRegisteredSerialNumber

                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SerialNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .SerialNumber))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CompanyID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .BrandID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SerialNumberConfirmed", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .SerialNumberConfirmed))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerOwnershipID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CustomerOwnershipID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RegisteredSerialNumberID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))



                End With

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "ID" back to the Structure
                sRegisteredSerialNumber.RegisteredSerialNumberID = SQLCommand.Parameters("@RegisteredSerialNumberID").Value

                If Not SQLCommand.Parameters("@RegisteredSerialNumberID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Registered Serial (Not System Error)"     'Set the Classes Error Message
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

#End Region
    End Class

    Public Class Knowledge

        '"Knowledge" Class Profile ******************************************************
        '
        '       Class: Knowledge 
        '       Parent File: WorkElements.vb
        '       Parent Project: CMS
        '       Parent Solution: CMS
        '       Created By: Amy Cook
        '       Created on: 07/21/05
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _lngKnowledgeID As Long          'Used to Hold the KnowledgeID currently being Used/Created
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs
#End Region

#Region "Class Properties"


        'Gets & Sets Current KnowledgeID
        Public Property KnowledgeID() As Long
            Get
                Return _lngKnowledgeID  'Gets Local _lngKnowledgeID Variable
            End Get
            Set(ByVal Value As Long)
                _lngKnowledgeID = Value 'Sets Local _lngKnowledgeID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
            Get
                Return _intErrorNumber
            End Get
        End Property


#End Region

#Region "Constructors"


        'Empty Consructor
        Public Sub New()

        End Sub

        'Use this Constructor when wanting to set the KnowledgeID on Creation
        Public Sub New(ByVal KnowledgeID As Integer)

            _lngKnowledgeID = KnowledgeID 'Sets Local KnowledgeID Variable

        End Sub




        ' Use this Constructor when wanting to create a NEW Product(Requires a "sKnowledge") 
        Public Sub New(ByRef sKnowledge As sKnowledge, ByRef CreatedKnowledge As Boolean)

            CreatedKnowledge = AddKnowledge(sKnowledge)

        End Sub




#End Region

#Region "Functions & Procedures"
        Public Function AddKnowledge(ByRef sKnowledge As sKnowledge) As Boolean


            'Function: AddKnowledge ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddKnowledge
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sKnowledge Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  08/22/05
            '       Description: Adds a Knowledge Entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "LLFBPS..[AddKnowledge]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Title", System.Data.SqlDbType.VarChar, 250))
                'SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.Text))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.Text, 999999999, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))




                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsFileOnly", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CanPublishExternal", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsQuickInfo", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@KnowledgeID").Value = sKnowledge.KnowledgeID
                SQLCommand.Parameters("@KnowledgeCategoryID").Value = sKnowledge.KnowledgeCategoryID
                SQLCommand.Parameters("@KnowledgeTypeID").Value = sKnowledge.KnowledgeTypeID
                SQLCommand.Parameters("@Title").Value = sKnowledge.Title
                SQLCommand.Parameters("@Description").Value = sKnowledge.Description
                SQLCommand.Parameters("@IsFileOnly").Value = sKnowledge.IsFileOnly
                SQLCommand.Parameters("@CreatedBy").Value = sKnowledge.CreatedBy
                SQLCommand.Parameters("@Active").Value = sKnowledge.IsActive
                SQLCommand.Parameters("@CanPublishExternal").Value = sKnowledge.CanPublishExternal
                SQLCommand.Parameters("@IsQuickInfo").Value = sKnowledge.IsQuickInfo

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "KnowledgeID" back to the sKnowledge Structure
                sKnowledge.KnowledgeID = SQLCommand.Parameters("@KnowledgeID").Value

                If Not SQLCommand.Parameters("@KnowledgeID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Knowledge(Not System Error)"     'Set the Classes Error Message
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
        Public Function AddKnowledgeDocument(ByVal sKnowledgeDocument As sKnowledgeDocument) As Boolean


            'Function: AddKnowledgeDocument ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddKnowledgeDocument
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sKnowledgeDocument Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  08/22/05
            '       Description: Adds a KnowledgeDocument Entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
                SQLCommand.CommandText = "LLFBPS..[AddKnowledgeDocument]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledgeDocument" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 75))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@LocalSystemPath", System.Data.SqlDbType.VarChar, 500))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeDocumentID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@Name").Value = sKnowledgeDocument.Name
                SQLCommand.Parameters("@LocalSystemPath").Value = sKnowledgeDocument.LocalSystemPath
                SQLCommand.Parameters("@CreatedBy").Value = sKnowledgeDocument.CreatedBy
                SQLCommand.Parameters("@CreatedName").Value = sKnowledgeDocument.CreatedByName
                SQLCommand.Parameters("@KnowledgeID").Value = sKnowledgeDocument.KnowledgeID


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "KnowledgeID" back to the sKnowledge Structure
                sKnowledgeDocument.KnowledgeDocumentID = SQLCommand.Parameters("@KnowledgeDocumentID").Value

                If Not SQLCommand.Parameters("@KnowledgeDocumentID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Knowledge Document(Not System Error)"     'Set the Classes Error Message
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
        Public Function AddKnowledgeResource(ByVal sKnowledgeResource As sKnowledgeResource) As Boolean

            'Function: AddKnowledgeResource ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddKnowledgeResource
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sKnowledgeResource Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  08/17/05
            '       Description: Adds a Knowledge Resource 
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Knowledge Resource One Step is Required: ********************
            '   - 1. Create a New Knowledge Resource Entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[AddKnowledgeResource]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledgeResource" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsAlert", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeResourceID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))

                SQLCommand.Parameters("@ResourceTypeID").Value = CType(sKnowledgeResource.ResourceTypeID, Integer)
                SQLCommand.Parameters("@ResourceID").Value = sKnowledgeResource.ResourceID
                SQLCommand.Parameters("@KnowledgeID").Value = sKnowledgeResource.KnowledgeID

                If sKnowledgeResource.IsAlert Then
                    SQLCommand.Parameters("@IsAlert").Value = 1
                Else
                    SQLCommand.Parameters("@IsAlert").Value = 0
                End If

                If sKnowledgeResource.Active Then
                    SQLCommand.Parameters("@Active").Value = 1
                Else
                    SQLCommand.Parameters("@Active").Value = 0
                End If


                SQLCommand.Parameters("@CreatedBy").Value = sKnowledgeResource.CreatedBy

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                'Pass the new "KnowledgeResourceID" back to the sKnowledgeResource Structure
                sKnowledgeResource.KnowledgeResourceID = SQLCommand.Parameters("@KnowledgeResourceID").Value

                'If there was a problem with Step One the New "KnowledgeResourceID" will NOT be greater then 0
                If Not SQLCommand.Parameters("@KnowledgeResourceID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Knowledge Resource (Not System Error)"     'Set the Classes Error Message
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
        Public Function AddKnowledgeKeyword(ByVal KnowledgeID As Long, ByVal KnowledgeKeyword As String, ByVal CreatedBy As Long) As Boolean

            'Function: AddKnowledgeResource ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddKnowledgeKeyword
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(2): KnowledgeID, KnowledgeKeyword, CreatedBy
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  08/23/05
            '       Description: Adds a Knowledge Keyword 
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
                SQLCommand.CommandText = "LLFBPS..[AddKnowledgeKeyword]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Keyword", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))


                SQLCommand.Parameters("@KnowledgeID").Value = KnowledgeID
                SQLCommand.Parameters("@Keyword").Value = KnowledgeKeyword
                SQLCommand.Parameters("@CreatedBy").Value = CreatedBy

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
        Public Function DeleteKnowledgeKeyword(ByVal KnowledgeKeywordID As Long) As Boolean

            'Function: DeleteKnowledgeKeyword ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteKnowledgeKeyword
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): KnowledgeKeywordID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  08/23/05
            '       Description: Deletes a Knowledge Keyword 
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
                SQLCommand.CommandText = "LLFBPS..[DeleteKnowledgeKeyword]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeKeywordID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@KnowledgeKeywordID").Value = KnowledgeKeywordID

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
        Public Function DeleteKnowledgeDocument(ByVal KnowledgeID As Integer, ByVal KnowledgeDocumentID As Integer) As Boolean

            'Function: DeleteKnowledgeDocument ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteKnowledgeDocument
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): KnowledgeDocumentID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  08/23/05
            '       Description: Deletes a Knowledge Document
            '       (If other knowledge is tied to this document, then only the KnowledgeDocumentJUNC will
            '       be deleted.  If not, the entry will be removed from the KnowledgeDocument table and
            '       entire document will be deleted)
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
                SQLCommand.CommandText = "LLFBPS..[DeleteKnowledgeDocument]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeDocumentID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@KnowledgeID").Value = KnowledgeID
                SQLCommand.Parameters("@KnowledgeDocumentID").Value = KnowledgeDocumentID

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
        Public Function DeleteKnowledgeResource(ByVal KnowledgeResourceID As Integer) As Boolean

            'Function: DeleteKnowledgeResource ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteKnowledgeResource
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): KnowledgeResourceID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  08/23/05
            '       Description: Deletes a KnowledgeResourceJUNC entry
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
                SQLCommand.CommandText = "LLFBPS..[DeleteKnowledgeResource]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeResourceID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@KnowledgeResourceID").Value = KnowledgeResourceID

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

        Public Function UpdateKnowledge(ByVal sKnowledge As sKnowledge) As Boolean

            'Function: UpdateKnowledge ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateKnowledge
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sKnowledge Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  08/22/05
            '       Description: Updates a Knowledge Entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Update Knowledge One Step is Required: ********************
            '   - 1. Update the Knowledge table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[UpdateKnowledge]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Title", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.Text))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsFileOnly", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CanPublishInternal", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CanPublishExternal", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsQuickInfo", System.Data.SqlDbType.Bit, 1))


                SQLCommand.Parameters("@KnowledgeID").Value = sKnowledge.KnowledgeID
                SQLCommand.Parameters("@KnowledgeCategoryID").Value = sKnowledge.KnowledgeCategoryID
                SQLCommand.Parameters("@KnowledgeTypeID").Value = sKnowledge.KnowledgeTypeID
                SQLCommand.Parameters("@Title").Value = sKnowledge.Title
                SQLCommand.Parameters("@Description").Value = sKnowledge.Description
                SQLCommand.Parameters("@IsFileOnly").Value = sKnowledge.IsFileOnly
                SQLCommand.Parameters("@Active").Value = sKnowledge.IsActive
                SQLCommand.Parameters("@CanPublishInternal").Value = sKnowledge.CanPublishInternal
                SQLCommand.Parameters("@CanPublishExternal").Value = sKnowledge.CanPublishExternal
                SQLCommand.Parameters("@IsQuickInfo").Value = sKnowledge.IsQuickInfo

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
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
        Public Function UpdateKnowledgeResource(ByVal sKnowledgeResource As sKnowledgeResource) As Boolean

            'Function: UpdateKnowledgeResource ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateKnowledgeResource
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: CMS
            '       Parent Solution: CMS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sKnowledgeResource Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  01/20/06
            '       Description: Updates a KnowledgeResourceJUNC Entry
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Update Knowledge One Step is Required: ********************
            '   - 1. Update the Knowledge table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[UpdateKnowledgeResourceJUNC]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values 

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsAlert", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))


                SQLCommand.Parameters("@KnowledgeResourceID").Value = sKnowledgeResource.KnowledgeResourceID

                If sKnowledgeResource.IsAlert Then
                    SQLCommand.Parameters("@IsAlert").Value = 1
                Else
                    SQLCommand.Parameters("@IsAlert").Value = 0
                End If

                If sKnowledgeResource.Active Then
                    SQLCommand.Parameters("@Active").Value = 1
                Else
                    SQLCommand.Parameters("@Active").Value = 0
                End If


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
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
        Public Function GetKnowledge(ByRef sKnowledge As sKnowledge) As Boolean

            'Function: GetKnowledge ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetKnowledge
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sKnowledge Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  07/20/05
            '       Description: Gets a piece of Knowledge
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If sKnowledge.KnowledgeID > 0 Then
                _lngKnowledgeID = sKnowledge.KnowledgeID
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
                SQLCommand.CommandText = "LLFBPS..[GetKnowledge]"              'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngKnowledgeID))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure
                If SQLDataReader.Read Then


                    With sKnowledge
                        'Fill Knowledge Information
                        If Not IsDBNull(SQLDataReader("KnowledgeID")) Then .KnowledgeID = SQLDataReader("KnowledgeID")
                        If Not IsDBNull(SQLDataReader("KnowledgeCategoryID")) Then .KnowledgeCategoryID = SQLDataReader("KnowledgeCategoryID")
                        If Not IsDBNull(SQLDataReader("KnowledgeTypeName")) Then .KnowledgeTypeName = SQLDataReader("KnowledgeTypeName")
                        If Not IsDBNull(SQLDataReader("KnowledgeTypeID")) Then .KnowledgeTypeID = SQLDataReader("KnowledgeTypeID")
                        If Not IsDBNull(SQLDataReader("KnowledgeCategoryName")) Then .KnowledgeCategoryName = SQLDataReader("KnowledgeCategoryName")
                        If Not IsDBNull(SQLDataReader("Title")) Then .Title = SQLDataReader("Title")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedByName")) Then .CreatedByName = SQLDataReader("CreatedByName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("Active")) Then .IsActive = SQLDataReader("Active")
                        If Not IsDBNull(SQLDataReader("IsFileOnly")) Then .IsFileOnly = SQLDataReader("IsFileOnly")
                        If Not IsDBNull(SQLDataReader("CanPublishInternal")) Then .CanPublishInternal = SQLDataReader("CanPublishInternal")
                        If Not IsDBNull(SQLDataReader("CanPublishExternal")) Then .CanPublishExternal = SQLDataReader("CanPublishExternal")
                        If Not IsDBNull(SQLDataReader("IsQuickInfo")) Then .IsQuickInfo = SQLDataReader("IsQuickInfo")

                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Knowledge Record Not Found"                   'Set the Classes Error Message
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
        Public Function GetKnowledgeDocument(ByRef sKnowledgeDocument As sKnowledgeDocument) As Boolean

            'Function: GetKnowledgeDocument ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetKnowledgeDocument
            '       Parent Class:  Knowledge
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sKnowledgeDocument Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  07/20/05
            '       Description: Gets a Knowledge Document
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If sKnowledgeDocument.KnowledgeID > 0 Then
                _lngKnowledgeID = sKnowledgeDocument.KnowledgeID
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
                SQLCommand.CommandText = "LLFBPS..[GetKnowledgeDocument]"              'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@KnowledgeID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngKnowledgeID))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure
                SQLDataReader.Read()
                If SQLDataReader.HasRows Then


                    With sKnowledgeDocument
                        'Fill Knowledge Information
                        If Not IsDBNull(SQLDataReader("KnowledgeDocumentID")) Then .KnowledgeDocumentID = SQLDataReader("KnowledgeDocumentID")
                        If Not IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
                        If Not IsDBNull(SQLDataReader("LocalSystemPath")) Then .LocalSystemPath = SQLDataReader("LocalSystemPath")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("KnowledgeDocumentJUNCID")) Then .KnowledgeDocumentJUNCID = SQLDataReader("KnowledgeDocumentJUNCID")
                    End With
                    bSuccess = True
                Else
                    bSuccess = False
                    _strErrorMessage = "Knowledge Document Not Found"                   'Set the Classes Error Message
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
        Public Property ProductSeriesID() As Long
            Get
                Return _lngProductSeriesID  'Gets Local _lngProductSeriesID Variable
            End Get
            Set(ByVal Value As Long)
                _lngProductSeriesID = Value 'Sets Local _lngProductID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
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
                SQLCommand.CommandText = "LLFBPS..[AddProductSeries]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 1000))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Freight", System.Data.SqlDbType.VarChar, 500))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSupported", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Specifications", System.Data.SqlDbType.VarChar, 1250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FirstCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SecondCategoryID", System.Data.SqlDbType.Int, 4))


                SQLCommand.Parameters("@Name").Value = sProductSeries.Name
                SQLCommand.Parameters("@Description").Value = sProductSeries.Description
                SQLCommand.Parameters("@Freight").Value = sProductSeries.Freight
                SQLCommand.Parameters("@IsSupported").Value = sProductSeries.IsSupported
                SQLCommand.Parameters("@Specifications").Value = sProductSeries.Specifications
                SQLCommand.Parameters("@PrimaryResourceImageID").Value = sProductSeries.PrimaryResourceImageID
                SQLCommand.Parameters("@CreatedBy").Value = sProductSeries.CreatedBy
                SQLCommand.Parameters("@CreatedName").Value = sProductSeries.CreatedByName
                SQLCommand.Parameters("@Active").Value = sProductSeries.Active

                SQLCommand.Parameters("@CompanyID").Value = sProductSeries.CompanyID
                SQLCommand.Parameters("@BrandID").Value = sProductSeries.BrandID
                SQLCommand.Parameters("@FirstCategoryID").Value = sProductSeries.ProductCategory1
                SQLCommand.Parameters("@SecondCategoryID").Value = sProductSeries.ProductCategory2


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "KnowledgeID" back to the sKnowledge Structure
                sProductSeries.ProductSeriesID = SQLCommand.Parameters("@ProductSeriesID").Value

                If Not SQLCommand.Parameters("@ProductSeriesID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product Series(Not System Error)"     'Set the Classes Error Message
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

        Public Function AddProductToSeries(ByVal sProduct As sProduct) As Boolean

            'Function: AddProductToSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductToSeries
            '       Parent Class:  ProductSeries
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
            '       Description: Adds a Product to a Series (ProductJUNC)
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add Product to a Series One Step is Required: ********************
            '   - 1. Add the ProductJUNC table entry
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[AddProductToSeries]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FirstCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SecondCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductJunctionID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@ProductID").Value = sProduct.ProductID
                SQLCommand.Parameters("@CompanyID").Value = sProduct.CompanyID
                SQLCommand.Parameters("@BrandID").Value = sProduct.BrandID
                SQLCommand.Parameters("@FirstCategoryID").Value = sProduct.ProductCategory1
                SQLCommand.Parameters("@SecondCategoryID").Value = sProduct.ProductCategory2
                SQLCommand.Parameters("@ProductSeriesID").Value = sProduct.ProductSeriesID
                SQLCommand.Parameters("@CreatedBy").Value = sProduct.CreatedBy
                SQLCommand.Parameters("@CreatedName").Value = sProduct.CreatedByName

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "KnowledgeID" back to the sKnowledge Structure
                sProduct.ProductJunctionID = SQLCommand.Parameters("@ProductJunctionID").Value

                If Not SQLCommand.Parameters("@ProductSeriesID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Product to the Series(Not System Error)"     'Set the Classes Error Message
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

        Public Function DeleteProductFromSeries(ByVal ProductJunctionID As Long) As Boolean

            'Function: DeleteProductFromSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteProductFromSeries
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductSeriesJunctionID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/17/05
            '       Description: Deletes a Product from a Series (ProductJUNC)
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
                SQLCommand.CommandText = "LLFBPS..[DeleteProductFromSeries]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductJunctionID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@ProductJunctionID").Value = ProductJunctionID

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

        Public Function DeleteProductSeriesRelationship(ByVal ProductSeriesJunctionID As Long) As Boolean

            'Function: DeleteProductSeriesRelationship ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteProductSeriesRelationship
            '       Parent Class:  ProductSeries
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): ProductSeriesJunctionID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  11/17/05
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
                SQLCommand.CommandText = "LLFBPS..[DeleteProductSeriesJunc]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesJunctionID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@ProductSeriesJunctionID").Value = ProductSeriesJunctionID

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
        Public Function AddProductSeriesRelationship(ByRef sProductSeries As sProductSeries) As Boolean

            'Function: AddProductSeries ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddProductSeriesRelationship
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
            '       Description: Adds a Product Series Relationship
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
                SQLCommand.CommandText = "LLFBPS..[AddProductSeriesJunc]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FirstCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SecondCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 75))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesJunctionID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))

                SQLCommand.Parameters("@ProductSeriesID").Value = sProductSeries.ProductSeriesID
                SQLCommand.Parameters("@CompanyID").Value = sProductSeries.CompanyID
                SQLCommand.Parameters("@BrandID").Value = sProductSeries.BrandID
                SQLCommand.Parameters("@FirstCategoryID").Value = sProductSeries.ProductCategory1
                SQLCommand.Parameters("@SecondCategoryID").Value = sProductSeries.ProductCategory2
                SQLCommand.Parameters("@CreatedBy").Value = sProductSeries.CreatedBy
                SQLCommand.Parameters("@CreatedName").Value = sProductSeries.CreatedByName

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
                        _strErrorMessage = "This series already belongs to this company." 'Set the Classes Error Message
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

            If sProductSeries.ProductSeriesID > 0 Then
                _lngProductSeriesID = sProductSeries.ProductSeriesID
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

                SQLCommand.CommandText = "LLFBPS..[GetProductSeries]"

                'Set the SeriesID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngProductSeriesID))


                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sProductSeries
                        'Fill Knowledge Information
                        If Not IsDBNull(SQLDataReader("ProductSeriesID")) Then .ProductSeriesID = SQLDataReader("ProductSeriesID")
                        If Not IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("Freight")) Then .Freight = SQLDataReader("Freight")
                        If Not IsDBNull(SQLDataReader("IsSupported")) Then .IsSupported = SQLDataReader("IsSupported")
                        If Not IsDBNull(SQLDataReader("Specifications")) Then .Specifications = SQLDataReader("Specifications")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")

                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID_Large")) Then .PrimaryResourceImageIDLarge = SQLDataReader("PrimaryResourceImageID_Large")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID_Small")) Then .PrimaryResourceImageIDSmall = SQLDataReader("PrimaryResourceImageID_Small")
                        If Not IsDBNull(SQLDataReader("PrimaryResourceImageID_feature")) Then .PrimaryResourceImageIDFeature = SQLDataReader("PrimaryResourceImageID_feature")
                        If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        If Not IsDBNull(SQLDataReader("ImagePath_large")) Then .ImagePathLarge = SQLDataReader("ImagePath_large")
                        If Not IsDBNull(SQLDataReader("ImagePath_small")) Then .ImagePathSmall = SQLDataReader("ImagePath_small")
                        If Not IsDBNull(SQLDataReader("ImagePath_feature")) Then .ImagePathFeature = SQLDataReader("ImagePath_feature")



                        'If Not IsDBNull(SQLDataReader("PrimaryResourceImageName")) Then .PrimaryResourceImageName = SQLDataReader("PrimaryResourceImageName")
                        If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("Active")) Then .Active = SQLDataReader("Active")
                        'If Not IsDBNull(SQLDataReader("CompanyID")) Then .CompanyID = SQLDataReader("CompanyID")
                        'If Not IsDBNull(SQLDataReader("CompanyName")) Then .CompanyName = SQLDataReader("CompanyName")
                        If Not IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
                        If Not IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
                        If Not IsDBNull(SQLDataReader("ParentCategoryID")) Then .ProductCategory1 = SQLDataReader("ParentCategoryID")
                        If Not IsDBNull(SQLDataReader("ParentCategoryName")) Then .ProductCategory1Name = SQLDataReader("ParentCategoryName")
                        If Not IsDBNull(SQLDataReader("ProductCategoryID")) Then .ProductCategory2 = SQLDataReader("ProductCategoryID")
                        If Not IsDBNull(SQLDataReader("CategoryName")) Then .ProductCategory2Name = SQLDataReader("CategoryName")

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
        Public Function UpdateProductSeriesImages(ByVal sProductSeries As sProductSeries) As Boolean

            'Function: UpdateProduct ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdateProductImages
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
            '       Created on:  08/27/07
            '       Description: Updates a Product's images
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
                SQLCommand.CommandText = "LLFBPS..[UpdateProductSeriesImages]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_large", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_small", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID_feature", System.Data.SqlDbType.Int, 4))


                SQLCommand.Parameters("@ProductSeriesID").Value = sProductSeries.ProductSeriesID
                SQLCommand.Parameters("@PrimaryResourceImageID").Value = sProductSeries.PrimaryResourceImageID
                SQLCommand.Parameters("@PrimaryResourceImageID_large").Value = sProductSeries.PrimaryResourceImageIDLarge
                SQLCommand.Parameters("@PrimaryResourceImageID_small").Value = sProductSeries.PrimaryResourceImageIDSmall
                SQLCommand.Parameters("@PrimaryResourceImageID_feature").Value = sProductSeries.PrimaryResourceImageIDFeature


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Updating Product Series Images (Not System Error)"     'Set the Classes Error Message
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
                SQLCommand.CommandText = "LLFBPS..[UpdateProductSeries]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 1000))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Freight", System.Data.SqlDbType.VarChar, 500))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSupported", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Specifications", System.Data.SqlDbType.VarChar, 1250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))


                SQLCommand.Parameters("@ProductSeriesID").Value = sProductSeries.ProductSeriesID
                SQLCommand.Parameters("@Name").Value = sProductSeries.Name
                SQLCommand.Parameters("@Description").Value = sProductSeries.Description
                SQLCommand.Parameters("@Freight").Value = sProductSeries.Freight
                SQLCommand.Parameters("@IsSupported").Value = sProductSeries.IsSupported
                If Not IsNothing(sProductSeries.Specifications) Then SQLCommand.Parameters("@Specifications").Value = sProductSeries.Specifications
                If Not IsNothing(sProductSeries.PrimaryResourceImageID) Then SQLCommand.Parameters("@PrimaryResourceImageID").Value = sProductSeries.PrimaryResourceImageID
                SQLCommand.Parameters("@Active").Value = sProductSeries.Active

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
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
        Public Property PartID() As Long
            Get
                Return _lngPartID  'Gets Local _lngProductID Variable
            End Get
            Set(ByVal Value As Long)
                _lngPartID = Value 'Sets Local _lngProductID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
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
        Public Function AddPartAlias(ByVal PartNumber As String, ByVal PartAlias As String) As Boolean

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
                SQLCommand.CommandText = "LLFBPS..[AddPartAlias]"     'Stored Procedure Name

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
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sPart Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  12/05/05
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
                SQLCommand.CommandText = "LLFBPS..[AddPart]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sPart" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedName", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BaseWarrantyDuration", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSellable", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))





                If sPart.MSRP > 0 Then
                    SQLCommand.Parameters("@MSRP").Value = sPart.MSRP
                Else
                    SQLCommand.Parameters("@MSRP").Value = DBNull.Value
                End If

                SQLCommand.Parameters("@IsSellable").Value = sPart.IsSellable
                SQLCommand.Parameters("@Name").Value = sPart.Name
                SQLCommand.Parameters("@Description").Value = sPart.Description
                SQLCommand.Parameters("@PartNumber").Value = sPart.PartNumber
                SQLCommand.Parameters("@PrimaryResourceImageID").Value = sPart.PrimaryResourceImageID
                SQLCommand.Parameters("@CreatedBy").Value = sPart.CreatedBy
                SQLCommand.Parameters("@CreatedName").Value = sPart.CreatedByName
                SQLCommand.Parameters("@BaseWarrantyDuration").Value = sPart.BaseWarrantyDuration
                SQLCommand.Parameters("@PartCategoryID").Value = sPart.PartCategoryID
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

        Public Function ApplyPartToProduct(ByVal PartID As Integer, ByVal ProductID As Integer) As Boolean

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
                SQLCommand.CommandText = "LLFBPS..[ApplyPartToProduct]"     'Stored Procedure Name

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
                SQLCommand.CommandText = "LLFBPS..[AddPartToProduct]"     'Stored Procedure Name

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
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Quantity", System.Data.SqlDbType.Int, 4))
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
                SQLCommand.Parameters("@Quantity").Value = sPart.ECommerce_Quantity
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
                SQLCommand.CommandText = "LLFBPS..[AddPartToProductSeries]"     'Stored Procedure Name

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
        Public Function UpdatePart(ByVal sPart As sPart) As Boolean

            'Function: UpdatePart ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: UpdatePart
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
            '       Description: Updates a Part 
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
                SQLCommand.CommandText = "LLFBPS..[UpdatePart]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PrimaryResourceImageID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BaseWarrantyDuration", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartCategoryID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MSRP", System.Data.SqlDbType.Money, 8, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsSellable", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, True, CType(18, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))

                SQLCommand.Parameters("@PartID").Value = sPart.PartID
                SQLCommand.Parameters("@Name").Value = sPart.Name
                SQLCommand.Parameters("@Description").Value = sPart.Description
                SQLCommand.Parameters("@PartNumber").Value = sPart.PartNumber
                SQLCommand.Parameters("@PrimaryResourceImageID").Value = sPart.PrimaryResourceImageID
                SQLCommand.Parameters("@BaseWarrantyDuration").Value = sPart.BaseWarrantyDuration
                SQLCommand.Parameters("@PartCategoryID").Value = sPart.PartCategoryID


                If sPart.MSRP > 0 Then
                    SQLCommand.Parameters("@MSRP").Value = sPart.MSRP
                Else
                    SQLCommand.Parameters("@MSRP").Value = DBNull.Value
                End If

                SQLCommand.Parameters("@IsSellable").Value = sPart.IsSellable


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
                SQLCommand.CommandText = "LLFBPS..[UpdateProductPartWarrantyDuration]"     'Stored Procedure Name

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
                SQLCommand.CommandText = "LLFBPS..[UpdateProductPart]"     'Stored Procedure Name

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
                SQLCommand.CommandText = "LLFBPS..[UpdateProductSeriesPart]"     'Stored Procedure Name

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
                SQLCommand.CommandText = "LLFBPS..[DeletePartProductSeriesJUNC]"     'Stored Procedure Name

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
                SQLCommand.CommandText = "LLFBPS..[DeletePartAlias]"     'Stored Procedure Name

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
                SQLCommand.CommandText = "LLFBPS..[DeletePartProductJUNC]"     'Stored Procedure Name

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

                SQLCommand.CommandText = "LLFBPS..[GetProductParts]"

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
                SQLCommand.CommandText = "LLFBPS..[GetPart]"

                'Set the ProductID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngPartID))

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
                        'If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
                        'If Not IsDBNull(SQLDataReader("PrimaryResourceImageName")) Then .PrimaryResourceImageName = SQLDataReader("PrimaryResourceImageName")
                        'If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                        'If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        'If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                        'If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        'If Not IsDBNull(SQLDataReader("BaseWarrantyDuration")) Then .BaseWarrantyDuration = SQLDataReader("BaseWarrantyDuration")
                        'If Not IsDBNull(SQLDataReader("PartCategoryID")) Then .PartCategoryID = SQLDataReader("PartCategoryID")
                        'If Not IsDBNull(SQLDataReader("PartCategoryName")) Then .PartCategoryName = SQLDataReader("PartCategoryName")
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
        Public Function GetPartByPartNumber(ByRef sPart As sPart) As Boolean

            'Function: GetPartByPartNumber ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: GetPartByPartNumber
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
            '       Created on:  03/28/08
            '       Description: Gets a Part by part number
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Try 'Only one "Try" statement 

                If sPart.PartNumber = "" Then
                    bSuccess = False
                    _strErrorMessage = "Part Number Not Specified"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                Else


                    SQLConn.Open() 'Open Database

                    'Set the Basic Command Information 
                    SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                    SQLCommand.Connection = SQLConn                          'Set the Connection
                    SQLCommand.CommandText = "LLFBPS..[GetPartByPartNumber]"

                    'Set the ProductID Parameter
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PartNumber", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sPart.PartNumber))

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
                            'If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
                            'If Not IsDBNull(SQLDataReader("PrimaryResourceImageName")) Then .PrimaryResourceImageName = SQLDataReader("PrimaryResourceImageName")
                            'If Not IsDBNull(SQLDataReader("ImagePath")) Then .ImagePath = SQLDataReader("ImagePath")
                            'If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                            'If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
                            'If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                            'If Not IsDBNull(SQLDataReader("BaseWarrantyDuration")) Then .BaseWarrantyDuration = SQLDataReader("BaseWarrantyDuration")
                            'If Not IsDBNull(SQLDataReader("PartCategoryID")) Then .PartCategoryID = SQLDataReader("PartCategoryID")
                            'If Not IsDBNull(SQLDataReader("PartCategoryName")) Then .PartCategoryName = SQLDataReader("PartCategoryName")
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

                End If

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
                SQLCommand.CommandText = "LLFBPS..[GetPartAliasNumber]"


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

    Public Class Customer

        '"Customer" Class Profile ******************************************************
        '
        '       Class:  Customer
        '       Parent File: WorkElements.vb
        '       Parent Project: CMS
        '       Parent Solution: CMS
        '       Created By: Amy Cook
        '       Created on: 07/21/05
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _lngCustomerID As Long          'Used to Hold the CustomerID currently being Used/Created
        Private _lngProductID As Long           'Used to Hold the ProductID currently being Used/Created 
        Private _lngCustomerOwnershipID As Long 'Used to Hold the CustomerOwnershipID currently being Used/Created
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs
        Private _dtCustomerListings As DataTable    'Used to hold customer listings




        Private Enum FindProductRegistrationType
            ByCustomerAndProduct = 1
            ByRegistrationID = 2
        End Enum


        Public Enum DuplicateEmailFoundStatus
            NotFound = 0
            CustomerFound = 1
            CustomerAndAccountFound = 2
            Unknown = 3
        End Enum

#End Region

#Region "Class Properties"


        'Gets & Sets Current CustomerID
        Public Property CustomerID() As Long
            Get
                Return _lngCustomerID  'Gets Local _lngCustomerID Variable
            End Get
            Set(ByVal Value As Long)
                _lngCustomerID = Value 'Sets Local _lngCustomerID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
            Get
                Return _intErrorNumber
            End Get
        End Property


        ' Returns Customer Listings from Data Table
        Public ReadOnly Property CustomerListings() As DataTable
            Get
                Return _dtCustomerListings
            End Get
        End Property


#End Region

#Region "Constructors"


        'Empty Consructor
        Public Sub New()

        End Sub

        'Use this Constructor when wanting to set the CustomerID on Creation
        Public Sub New(ByVal CustomerID As Integer)

            _lngCustomerID = CustomerID 'Sets Local CustomerID Variable

        End Sub




        ' Use this Constructor when wanting to create a NEW Customer(Requires a "sCustomer") 
        Public Sub New(ByRef sCustomer As sCustomer, ByRef CreatedCustomer As Boolean)

            CreatedCustomer = AddCustomer(sCustomer)

        End Sub




#End Region

#Region "Functions & Procedures"


        Public Function ValidateTempCustomerAccount(ByVal CustomerID As Integer, ByVal CustomerKey As String) As Boolean

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            'Dim SQLTran As SqlTransaction = Nothing
            ' SQLTransaction Object - Used to guarantee that ALL or None of the function Processes  

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Update a Customer  Two Steps are Required: ********************
            '   - 1. Update the Customers Information
            '   - 2. Create a New Customer History appended to the Customer 
            '   
            '*********************************************************************************************

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database
                '     SQLTran = SQLConn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin your Transaction


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection
                '    SQLCommand.Transaction = SQLTran                     'Set the Transaction



                'Step One - Update the Customers Information 

                'Set the Specific Command Information for Step One
                SQLCommand.CommandText = "LLFWebsite..[ValidateTempCustomerIDandKey]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@tempCustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerKey", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerKey))

                bSuccess = SQLCommand.ExecuteScalar() 'Execute the Stored Procedure


                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                '   SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts


                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                '     SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function

        Public Function AddCustomer(ByRef sCustomer As sCustomer) As Boolean

            'Function: AddCustomer ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddCustomer
            '       Parent Class:  Customer
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sCustomer Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  12/05/05
            '       Description: Adds a Customer
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLTran As SqlTransaction = Nothing
            ' SQLTransaction Object - Used to guarantee that ALL or None of the function Processes  

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add a Customer  Two Steps are Required: ********************
            '   - 1. Create a New Customer  
            '   - 2. Create a New Customer History appended to the new Customer 
            '   
            '*********************************************************************************************

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database
                SQLTran = SQLConn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin your Transaction


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection
                SQLCommand.Transaction = SQLTran                     'Set the Transaction



                'Step One - Create a New Customer  

                'Set the Specific Command Information for Step One
                SQLCommand.CommandText = "LLFBPS..[AddCustomer]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FirstName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.FirstName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MiddleName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.MiddleName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@LastName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.LastName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Gender", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Gender))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address1", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Address1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address2", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Address2))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@City", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.City))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@State", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.State))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Zip", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Zip))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Country", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Country))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Email", System.Data.SqlDbType.VarChar, 250, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Email))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@HomePhone", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.HomePhone))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WorkPhone", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.WorkPhone))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MobilePhone", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.MobilePhone))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FaxNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.FaxNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AlternateCustomerNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.AlternateCustomerNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AlternateCustomerNumber2", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.AlternateCustomerNumber2))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField1", System.Data.SqlDbType.VarChar, 1500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.UserField1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField2", System.Data.SqlDbType.VarChar, 1500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.UserField2))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField3", System.Data.SqlDbType.VarChar, 1500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.UserField3))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.CreatedBy))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Now))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Active))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsBusiness", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.IsBusiness))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TaxExemptNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.TaxExemptNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RetailerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.RetailerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.CustomerID))


                SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure

                'Pass the new "CustomerID" back to the sCustomer Structure
                sCustomer.CustomerID = SQLCommand.Parameters("@CustomerID").Value


                'If there was a problem with Step One the New "CustomerID" will NOT be greater then 0
                If sCustomer.CustomerID > 0 Then 'IF True Continue to Step Two


                    'Set the Specific Command Information for Step Two
                    SQLCommand.Parameters.Clear() 'Clear the Parameters 
                    SQLCommand.CommandText = "LLFBPS..[AddCustomerHistory]" 'Stored Procedure Name

                    'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure & Private Variables ( m_lngCustomerID)
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomer.CustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, String.Format("Customer # {0} was created", sCustomer.CustomerID)))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomer.CreatedBy))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Now))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerHistoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                    SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure



                    'If there was a problem with Step Two the New "CustomerHistoryID" will NOT be greater then 0

                    If Not SQLCommand.Parameters("@CustomerHistoryID").Value > 0 Then 'If True then if FAILED Step Three

                        _strErrorMessage = "Error adding Customer History"             'Set the Classes Error Message
                        _intErrorNumber = 0                                        'Set the Classes Error Number
                        bSuccess = False                                          'Set the "Succeed" Boolean to False
                        SQLTran.Rollback()                                          'RollBack all SQLCommand Inserts

                    End If



                Else 'Failed Step One - Adding the New Customer 


                    _strErrorMessage = "Internal Error adding Customer (Not System Error)"  'Set the Classes Error Message
                    _intErrorNumber = 0         'Set the Classes Error Number
                    bSuccess = False           'Set the "Succeed" Boolean to False
                    SQLTran.Rollback()           'RollBack all SQLCommand Inserts


                End If

                '*************************************************************************************************
                'IF BOTH STEPS processed successfully then COMMIT the SQLCOmmand INSERTS
                If bSuccess = True Then SQLTran.Commit()
                '*************************************************************************************************

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts


                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function



        Public Function UpdateCustomer(ByRef sCustomer As sCustomer) As Boolean

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLTran As SqlTransaction = Nothing
            ' SQLTransaction Object - Used to guarantee that ALL or None of the function Processes  

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Update a Customer  Two Steps are Required: ********************
            '   - 1. Update the Customers Information
            '   - 2. Create a New Customer History appended to the Customer 
            '   
            '*********************************************************************************************

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database
                SQLTran = SQLConn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin your Transaction


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection
                SQLCommand.Transaction = SQLTran                     'Set the Transaction



                'Step One - Update the Customers Information 

                'Set the Specific Command Information for Step One
                SQLCommand.CommandText = "LLFBPS..[UpdateCustomer]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FirstName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.FirstName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MiddleName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.MiddleName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@LastName", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.LastName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Gender", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Gender))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address1", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Address1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address2", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Address2))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@City", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.City))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@State", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.State))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Zip", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Zip))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Country", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Country))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Email", System.Data.SqlDbType.VarChar, 250, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Email))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@HomePhone", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.HomePhone))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@WorkPhone", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.WorkPhone))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MobilePhone", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.MobilePhone))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FaxNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.FaxNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AlternateCustomerNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.AlternateCustomerNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AlternateCustomerNumber2", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.AlternateCustomerNumber2))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField1", System.Data.SqlDbType.VarChar, 1500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.UserField1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField2", System.Data.SqlDbType.VarChar, 1500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.UserField2))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserField3", System.Data.SqlDbType.VarChar, 1500, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.UserField3))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.CreatedBy))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Now))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.Active))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsBusiness", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.IsBusiness))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TaxExemptNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.TaxExemptNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RetailerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomer.RetailerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Errors", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, 1))


                SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure


                'If there was a problem with Step One the @Errors Parameters will be greater then 0

                If SQLCommand.Parameters("@Errors").Value = 0 Then  'IF True Continue to Step Two


                    'Set the Specific Command Information for Step Two
                    SQLCommand.Parameters.Clear() 'Clear the Parameters 
                    SQLCommand.CommandText = "LLFBPS..[AddCustomerHistory]" 'Stored Procedure Name

                    'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure & Private Variables ( m_lngCustomerID)
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomer.CustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomer._LogInformation))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomer.CreatedBy))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Now))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerHistoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                    SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure



                    'If there was a problem with Step Two the New "CustomerHistoryID" will NOT be greater then 0

                    If Not SQLCommand.Parameters("@CustomerHistoryID").Value > 0 Then 'If True then if FAILED Step Three

                        _strErrorMessage = "Error adding Customer History"             'Set the Classes Error Message
                        _intErrorNumber = 0                                        'Set the Classes Error Number
                        bSuccess = False                                          'Set the "Succeed" Boolean to False
                        SQLTran.Rollback()                                          'RollBack all SQLCommand Inserts

                    End If



                Else 'Failed Step One - Updating the  Customer 


                    _strErrorMessage = "Internal Error updating Customer (Not System Error)"  'Set the Classes Error Message
                    _intErrorNumber = 0         'Set the Classes Error Number
                    bSuccess = False           'Set the "Succeed" Boolean to False
                    SQLTran.Rollback()           'RollBack all SQLCommand Inserts


                End If

                '*************************************************************************************************
                'IF BOTH STEPS processed successfully then COMMIT the SQLCOmmand INSERTS
                If bSuccess = True Then SQLTran.Commit()
                '*************************************************************************************************

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts


                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function


        Public Function CustomerMatch(ByVal sCustomer As sCustomer, Optional ByRef Confirmed As Boolean = False, Optional ByRef ResultInformation As String = "") As Boolean



            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed
            Dim strFirstName As String = ""
            Dim strLastName As String = ""
            Dim strAddress As String = ""
            Dim strZip As String = ""
            Dim strHomePhone As String = ""
            Dim strEmail As String = ""
            Dim intMatchCode As Int16 = 0
            Dim strTempMessage As String = ""
            Dim blnContinue As Boolean = True

            ResultInformation = "" 'Preset to nothing
            Confirmed = False  'Preset this to False


            'DB Objects
            Dim dtCustomers As DataTable 'Customer Data Table 
            Dim drCustomers As DataRow   'Customer Data Row
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object
            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




            Try 'Only one "Try" statement 

                'Create the Data Table
                dtCustomers = New DataTable("Customers")
                dtCustomers.Columns.Add("CustomerID", GetType(Long))
                '   dtCustomers.Columns.Add("Firstname", GetType(String))
                '   dtCustomers.Columns.Add("Middlename", GetType(String))
                '   dtCustomers.Columns.Add("Lastname", GetType(String))
                dtCustomers.Columns.Add("Name", GetType(String))
                dtCustomers.Columns.Add("Address", GetType(String))
                '  dtCustomers.Columns.Add("Address2", GetType(String))
                dtCustomers.Columns.Add("City", GetType(String))
                dtCustomers.Columns.Add("State", GetType(String))
                dtCustomers.Columns.Add("Zip", GetType(String))
                dtCustomers.Columns.Add("HomePhone", GetType(String))
                dtCustomers.Columns.Add("MobilePhone", GetType(String))
                dtCustomers.Columns.Add("Email", GetType(String))




                SQLConn.Open() 'Open Database

                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                          'Set the Connection



                'First Check to see if you have a Name/Email Match
                SQLCommand.CommandText = "LLFBPS..[GetCustomerByMatchCodes]"              'Stored Procedure Name
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Firstname", System.Data.SqlDbType.VarChar, 150, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Lastname", System.Data.SqlDbType.VarChar, 150, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Address", System.Data.SqlDbType.VarChar, 150, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Zip", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@HomePhone", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Email", System.Data.SqlDbType.VarChar, 250, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MatchCodeType", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Nothing))



                Do Until intMatchCode > 19 OrElse blnContinue = False


                    With sCustomer

                        Select Case intMatchCode

                            Case 0
                                strFirstName = .FirstName
                                strLastName = .LastName
                                strAddress = .Address1
                                strZip = .Zip
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the First name, Last name, Address and Zip"

                            Case 1
                                strFirstName = .FirstName
                                strLastName = .LastName
                                strAddress = .Address1
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the First name, Last name, and Address"


                            Case 2
                                strFirstName = .FirstName
                                strLastName = .LastName
                                If Len(.Email) > 4 Then
                                    strEmail = .Email
                                Else
                                    strEmail = ""
                                End If
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the First name, Last name, and Email address"

                            Case 3
                                strFirstName = .FirstName
                                strLastName = .LastName
                                If Len(.HomePhone) > 6 Then
                                    strHomePhone = .HomePhone
                                Else
                                    strHomePhone = ""
                                End If
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the First name, Last name, and Home Phone"
                            Case 4
                                strLastName = .LastName
                                strAddress = .Address1
                                strZip = .Zip
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the Last name, Address and Zip"

                            Case 5
                                strLastName = .LastName
                                If Len(.HomePhone) > 6 Then
                                    strHomePhone = .HomePhone
                                Else
                                    strHomePhone = ""
                                End If
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the Lastname and Home Phone"


                            Case 6
                                If Not .FirstName = "" Then strFirstName = Trim(Left(.FirstName, 2)) & "%"
                                If Not .LastName = "" Then strLastName = Trim(Left(.LastName, 2)) & "%"
                                If Len(.Email) > 4 Then
                                    strEmail = .Email
                                Else
                                    strEmail = ""
                                End If
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching Potential First & Last name, and Email address"


                            Case 7
                                If Not .FirstName = "" Then strFirstName = Trim(Left(.FirstName, 2)) & "%"
                                If Not .LastName = "" Then strLastName = Trim(Left(.LastName, 2)) & "%"
                                If Len(.HomePhone) > 6 Then
                                    strHomePhone = .HomePhone
                                Else
                                    strHomePhone = ""
                                End If

                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching Potential First & Last name, and Home phone"

                            Case 8
                                strAddress = .Address1
                                strZip = .Zip
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the Address and Zip"


                            Case 9
                                If Len(.Email) > 4 Then
                                    strEmail = .Email
                                Else
                                    strEmail = ""
                                End If
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the Email address"


                            Case 10

                                If Len(.HomePhone) > 6 Then
                                    strHomePhone = .HomePhone
                                Else
                                    strHomePhone = ""
                                End If
                                Confirmed = True 'This is a confirmed match
                                strTempMessage = "Results determined by matching the Home phone"


                            Case 11
                                strFirstName = .FirstName
                                strLastName = .LastName
                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching the First & Last name"

                            Case 12
                                strLastName = .LastName
                                strZip = .Zip
                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching the Lastname & Zip"


                            Case 13
                                If Not .FirstName = "" Then strFirstName = Trim(Left(.FirstName, 1)) & "%"
                                If Not .LastName = "" Then strLastName = Trim(Left(.LastName, 1)) & "%"
                                strZip = .Zip
                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching Potential First & Last name, and Zip"


                            Case 14
                                If Not .FirstName = "" Then strFirstName = Trim(Left(.FirstName, 1)) & "%"
                                strLastName = .LastName
                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching Potential Names"


                            Case 15
                                strFirstName = .FirstName
                                If Not .LastName = "" Then strLastName = Trim(Left(.LastName, 1)) & "%"
                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching Potential Names"

                            Case 16
                                If Not .FirstName = "" Then strFirstName = String.Format("{0}%{1}", Trim(Left(.FirstName, 1)), Trim(Right(.FirstName, 1)))
                                If Not .LastName = "" Then strLastName = String.Format("{0}%{1}", Trim(Left(.LastName, 1)), Trim(Right(.LastName, 1)))
                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching Potential Names"

                            Case 17
                                If .Zip <> "" Then
                                    strZip = .Zip
                                Else
                                    strZip = "00099000"
                                End If

                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching Zip Codes"

                            Case 18
                                strLastName = .LastName
                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching Last Name"

                            Case 19
                                strFirstName = .FirstName
                                Confirmed = False  'This is NOT a confirmed match
                                strTempMessage = "Results determined by matching First Name"


                        End Select

                    End With

                    'Set the parameters values
                    SQLCommand.Parameters("@Firstname").Value = strFirstName
                    SQLCommand.Parameters("@Lastname").Value = strLastName
                    SQLCommand.Parameters("@Address").Value = strAddress
                    SQLCommand.Parameters("@Zip").Value = strZip
                    SQLCommand.Parameters("@HomePhone").Value = strHomePhone
                    SQLCommand.Parameters("@Email").Value = strEmail
                    SQLCommand.Parameters("@MatchCodeType").Value = intMatchCode

                    'Set the DataReader
                    SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                    Do While SQLDataReader.Read()

                        blnContinue = False 'No need to continue


                        drCustomers = dtCustomers.NewRow
                        drCustomers("CustomerID") = SQLDataReader("CustomerID")
                        '  drCustomers("Firstname") = SQLDataReader("Firstname")
                        '  drCustomers("Middlename") = SQLDataReader("Middlename")
                        '  drCustomers("Lastname") = SQLDataReader("Lastname")
                        drCustomers("Name") = String.Format("{0} {1} {2}", SQLDataReader("Firstname"), SQLDataReader("Middlename"), SQLDataReader("lastname"))
                        drCustomers("Address") = String.Format("{0} {1}", SQLDataReader("Address1"), SQLDataReader("Address2"))
                        'drCustomers("Address2") = SQLDataReader("Address2")
                        drCustomers("City") = SQLDataReader("City")
                        drCustomers("State") = SQLDataReader("State")
                        drCustomers("Zip") = SQLDataReader("Zip")
                        drCustomers("HomePhone") = SQLDataReader("HomePhone")
                        drCustomers("MobilePhone") = SQLDataReader("MobilePhone")
                        drCustomers("Email") = SQLDataReader("Email")
                        dtCustomers.Rows.Add(drCustomers)

                    Loop

                    'Close Data Reader
                    SQLDataReader.Close()


                    If blnContinue = False Then
                        'Set the local "Customer Listings Property if Results are found
                        _dtCustomerListings = dtCustomers.Copy

                    End If

                    intMatchCode += 1

                Loop




                'Close Connection
                SQLConn.Close()


                'Prepare the Return "Match" message based off the return information
                If blnContinue = True Then 'if this is true, then NO results were found
                    Confirmed = False 'Preset back to false
                    ResultInformation = "No Customer Matches Found"

                Else ' A match was found, now check to see if it was confirmed

                    If Confirmed = True Then
                        ResultInformation = String.Format("CONFIRMED matches found - {0}", strTempMessage)
                    Else
                        ResultInformation = String.Format("Potential matches found - {0}", strTempMessage)
                    End If

                End If



                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False

                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number


                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number



            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                dtCustomers = Nothing
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing


            End Try

            'Set returning Boolean

            If blnContinue = True Then
                Return False
            Else
                Return True
            End If

        End Function
        Public Function DeleteCustomerProductOwnership(ByVal CustomerOwnerShipID As Long) As Boolean

            'Function: DeleteCustomerProductOwnership ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: DeleteCustomerProductOwnership
            '       Parent Class:  Customer
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): CustomerOwnerShipID
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  07/05/06
            '       Description: Deletes a Customer Registration entry (Customer Product Ownership)
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
                SQLCommand.CommandText = "LLFBPS..[DeleteCustomerProductOwnership]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sShipToAddress" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerOwnershipID", System.Data.SqlDbType.Int, 4))

                SQLCommand.Parameters("@CustomerOwnershipID").Value = CustomerOwnerShipID

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
        Public Function GetCustomer(ByRef sCustomer As sCustomer) As Boolean

            Dim blnsuccess As Boolean = False

            'Set the CustomerID
            If sCustomer.CustomerID > 0 Then
                _lngCustomerID = sCustomer.CustomerID
            End If

            If _lngCustomerID > 0 Then
                blnsuccess = _GetCustomer(sCustomer)
            End If

            Return blnsuccess

        End Function

        Public Function GetCustomer(ByRef sCustomer As sCustomer, ByVal CustomerID As Long) As Boolean

            Dim blnsuccess As Boolean = False

            'Set the CustomerID
            _lngCustomerID = CustomerID

            If _lngCustomerID > 0 Then
                blnsuccess = _GetCustomer(sCustomer)
            End If

            Return blnsuccess
        End Function
        Private Function GeneratePassword(ByVal passwordLength As Integer) As String
            Dim Vowels() As Char = New Char() {"a", "e", "i", "o", "u"}
            Dim Consonants() As Char = New Char() {"b", "c", "d", "f", "g", "h", "j", "k", "l", "m", "n", "p", "r", "s", "t", "v"}
            Dim DoubleConsonants() As Char = New Char() {"c", "d", "f", "g", "l", "m", "n", "p", "r", "s", "t"}
            Dim wroteConsonant As Boolean  'boolean   
            Dim counter As Integer
            Dim rnd As New Random
            Dim passwordBuffer As New StringBuilder
            wroteConsonant = False
            For counter = 0 To passwordLength
                If passwordBuffer.Length > 0 And (wroteConsonant = False) And (rnd.Next(100) < 10) Then
                    passwordBuffer.Append(DoubleConsonants(rnd.Next(DoubleConsonants.Length)), 2)
                    counter += 1
                    wroteConsonant = True
                Else
                    If (wroteConsonant = False) And (rnd.Next(100) < 90) Then
                        passwordBuffer.Append(Consonants(rnd.Next(Consonants.Length)))
                        wroteConsonant = True
                    Else
                        passwordBuffer.Append(Vowels(rnd.Next(Vowels.Length)))
                        wroteConsonant = False
                    End If
                End If
            Next
            'size the buffer    
            passwordBuffer.Length = passwordLength
            Return passwordBuffer.ToString
        End Function


        Public Function CheckForDuplicateEmail(ByVal Email As String, Optional ByVal AccountID As Integer = 0) As DuplicateEmailFoundStatus



            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim DuplicateEmailFoundStatus As DuplicateEmailFoundStatus = DuplicateEmailFoundStatus.NotFound

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            Dim SQLDataReader As SqlDataReader

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[GetCustomerByEmail]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Email", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Email.Trim))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


                If SQLDataReader.Read = True Then
                    DuplicateEmailFoundStatus = DuplicateEmailFoundStatus.CustomerFound
                End If


                If AccountID > 0 Then 'Check to see if an Accounts' been open on the Email
                    SQLDataReader.Close()
                    SQLCommand.Parameters.Clear()


                    'Set the Specific Command Information 
                    SQLCommand.CommandText = "LLFBPS..[GetAccountByUserName]"  'Stored Procedure Name

                    'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Username", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Email.Trim))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AccountID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, AccountID))

                    'Set the DataReader
                    SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)
                    If SQLDataReader.Read = True Then
                        DuplicateEmailFoundStatus = DuplicateEmailFoundStatus.CustomerAndAccountFound
                    End If


                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                DuplicateEmailFoundStatus = DuplicateEmailFoundStatus.Unknown
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                DuplicateEmailFoundStatus = DuplicateEmailFoundStatus.Unknown

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return DuplicateEmailFoundStatus
        End Function
        Public Function GetCustomerNumberByEmail(ByVal Email As String) As Integer

            'This will return THE FIRST customer number found that matches the email address

            Dim CustomerID As Integer

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            Dim SQLDataReader As SqlDataReader

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[GetCustomerByEmail]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Email", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Email.Trim))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


                If SQLDataReader.Read = True Then
                    CustomerID = SQLDataReader("CustomerID")
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                CustomerID = 0

                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                CustomerID = 0


                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return CustomerID


        End Function
        Public Overloads Function ResetCustomerAccountPassword(ByVal CustomerID As Integer, ByVal AccountID As Integer, ByVal strPassword As String) As Boolean

            Dim blnsuccess As Boolean = False
            Dim strHashedPassword As String

            'Set the CustomerID
            _lngCustomerID = CustomerID

            strHashedPassword = FormsAuthentication.HashPasswordForStoringInConfigFile(strPassword, ConfigurationManager.AppSettings("MyCharbroilKey").ToString)
            If _lngCustomerID > 0 Then
                blnsuccess = _ResetCustomerPassword(_lngCustomerID, AccountID, strHashedPassword)
            End If

            Return blnsuccess

        End Function
        Public Function GetCustomerIDByAccountEmail(ByVal CustomerEmailAddress As String, ByVal AccountID As Integer) As Integer


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim CustomerID As Integer = 0

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            Dim SQLDataReader As SqlDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[GetCustomerAccount]"  'Stored Procedure Name


                'Stored Procedure Paramaters 
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Username", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CustomerEmailAddress))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AccountID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, AccountID))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


                If SQLDataReader.Read = True Then
                    _lngCustomerID = SQLDataReader("CustomerID")
                    If _lngCustomerID > 0 Then
                        CustomerID = _lngCustomerID
                    End If
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                CustomerID = 0
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                CustomerID = 0
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return CustomerID



        End Function
        Public Function ResetCustomerAccountUsername(ByVal CustomerID As Integer, ByVal AccountID As Integer, ByVal Username As String) As Boolean
            Dim bSuccess As Boolean 'Return variable 


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object

            Try


                'Update Account password

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[UpdateCustomerAccountUsername]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCaseActionItem" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AccountID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Username", System.Data.SqlDbType.VarChar, 150))

                SQLCommand.Parameters("@AccountID").Value = AccountID
                SQLCommand.Parameters("@CustomerID").Value = CustomerID
                SQLCommand.Parameters("@Password").Value = Username

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Resetting Customer Account Username (Not System Error)"     'Set the Classes Error Message
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
        Private Function _ResetCustomerPassword(ByVal CustomerID As Integer, ByVal AccountID As Integer, ByVal strHashedPassword As String) As Boolean
            Dim bSuccess As Boolean 'Return variable 


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object

            Try

                'Update Account password

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[UpdateCustomerAccountPassword]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCaseActionItem" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AccountID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Password", System.Data.SqlDbType.VarChar, 150))

                SQLCommand.Parameters("@AccountID").Value = AccountID
                SQLCommand.Parameters("@CustomerID").Value = CustomerID
                SQLCommand.Parameters("@Password").Value = strHashedPassword

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                If SQLCommand.Parameters("@RETURN_VALUE").Value = 0 Then
                    bSuccess = True
                Else
                    _intErrorNumber = "Internal Error Resetting Customer Account Password (Not System Error)"     'Set the Classes Error Message
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
        Private Function EncryptQueryStringValue(ByVal QueryStringValue As String) As String
            'This Encrypts the value using the QueryStringKey and the persons Identity
            Return System.Web.Security.FormsAuthentication.HashPasswordForStoringInConfigFile(QueryStringValue, "sha1")

        End Function
        Private Function SendRegistrationConfirmationEmail(ByVal sEmail As sEmail) As Boolean

            'Sends a custom email based on the sEmail structure
            Dim Email As New MailMessage
            Dim blnSentEmail As Boolean = True
            Dim strStandardHeader As String = ""
            Dim strBodyHeader As String = ""
            Dim strReplyInfoHeader As String = ""
            Dim strReplyInfoFooter As String = ""
            Dim strRedirectReplyLink As String = ""
            Dim strCreateAccountLink As String = "" 'Link that allows customer to create account
            Dim sAccount As BPS_BL.BPS.sAccount = Nothing

            Dim SmtpMail As New SmtpClient
            Dim strImagePath As String

            strImagePath = ConfigurationManager.AppSettings("ApplicationHome") & "/images/"

            Try
                With sEmail
                    If .ReplyTo <> "" Then
                        Email.Headers.Add("Reply-To", .ReplyTo)
                    End If

                    'Build Standard Header to be used in all Customer Emails
                    strStandardHeader = "<html><style type=""text/css"">"
                    strStandardHeader = strStandardHeader & "<!--.main {font-family: Arial, Helvetica, sans-serif;font-size: 14px;font-weight: normal;}--></style>" & vbCrLf
                    strStandardHeader = strStandardHeader & "<body><table width=""766"" border=""0"" cellpadding=""0"">" & vbCrLf
                    strStandardHeader = strStandardHeader & "<tr><td><img src=""" & strImagePath & "topheader.gif"" width=""766""  border=""0"" " & vbCrLf
                    strStandardHeader = strStandardHeader & "alt="""" usemap=""#header_Map""></td></tr>" & vbCrLf

                    'Build Standard Footer (Non-Reply)
                    strReplyInfoFooter = "</td></tr><tr><td><img src=""" & strImagePath & "footerb.gif"" width=""766"" height=""32"" border=""0"" alt=""""></td></tr>" & vbCrLf
                    strReplyInfoFooter = strReplyInfoFooter & "</table><map name=""header_Map""><area shape=""rect"" alt=""Char-Broil Website"" coords=""0,0,173,47"" href=""" & ConfigurationManager.AppSettings("ApplicationHome") & """>" & vbCrLf
                    strReplyInfoFooter = strReplyInfoFooter & "<area shape=""rect"" alt=""Respond to this email"" coords=""343,56,382,73"" href=""" & strRedirectReplyLink & """></map>" & vbCrLf
                    strReplyInfoFooter = strReplyInfoFooter & "<map name=""footer_Map""><area shape=""rect"" alt=""Respond to this email"" coords=""582,4,739,23"" href=""" & strRedirectReplyLink & """></map></body></html>" & vbCrLf


                    If sEmail.CustomerID > 0 Then
                        sAccount.CustomerID = sEmail.CustomerID
                        sAccount.AccountID = CType(ConfigurationManager.AppSettings("MyCharbroilAccountID").ToString, Integer)

                        If Not GetAccount(sAccount, CustomerAccountAuthenticationMode.CustomerAndAccountID) Then
                            strCreateAccountLink = "<br><br><a href=""" & String.Format(ConfigurationManager.AppSettings("ApplicationHome") & "/CustomerHomeWebsite/CustomerHomePasswordUpdate.aspx?source={0}&CustomerID={1}&CustomerIDKey={2}&AccountID={3}&AccountIDKey={4}", "email", sEmail.CustomerID, EncryptQueryStringValue(sEmail.CustomerID), CType(ConfigurationManager.AppSettings("MyCharbroilAccountID").ToString, Integer), EncryptQueryStringValue(CType(ConfigurationManager.AppSettings("MyCharbroilAccountID").ToString, Integer))) & """>Get access to the latest information for your grill by creating an account.  You will have access to product manuals, part listings, maintenance and troubleshooting help and lots more!</a>" & vbCrLf
                        End If

                    End If
                    strBodyHeader = "<tr><td align=center><table width=""735"" border=""0"" cellpadding=""5""><tr><td align=left valign=top class=main><br>" & vbCrLf


                    If InStr(sEmail.MailToAddress, "@") Then
                        Email.Priority = .Priority
                        Email.To.Add(New MailAddress(.MailToAddress))
                        Email.From = New MailAddress(.MailFromAddress)
                        Email.Subject = .Subject
                        Email.IsBodyHtml = True
                        Email.BodyEncoding = System.Text.Encoding.UTF8
                        Email.Body = strStandardHeader & strReplyInfoHeader & strBodyHeader & .Body & strCreateAccountLink & strReplyInfoFooter

                        Dim strAttachment As String
                        If Not IsNothing(sEmail.Attachments) Then
                            For Each strAttachment In sEmail.Attachments
                                If strAttachment <> "" Then
                                    Dim delim As Char = ","
                                    Dim sSubstr As String
                                    For Each sSubstr In strAttachment.Split(delim)
                                        Dim myAttachment As Net.Mail.Attachment = New Net.Mail.Attachment(sSubstr)
                                        Email.Attachments.Add(myAttachment)
                                    Next
                                End If

                            Next
                        End If
                        SmtpMail.Host = .SMTPServer
                        'If AllowOutboundEmail, then send the email,
                        'Otherwise skip this function
                        If ConfigurationManager.AppSettings("AllowOutboundEmail").ToString = True Then
                            SmtpMail.Send(Email)
                        End If
                    End If
                    blnSentEmail = True
                    Email = Nothing

                End With
            Catch ex As Exception
                _strErrorMessage = ex.Message.ToString       'Set the Classes Error Message
                _intErrorNumber = 0       'Set the Classes Error Number
                blnSentEmail = False
            Finally
                SendRegistrationConfirmationEmail = blnSentEmail

            End Try
        End Function

        'Public Function SendCustomerPasswordChangeEmail(ByVal CustomerEmail As String, ByVal AccountID As Integer) As Boolean
        '    Dim sEmail As sEmail
        '    Dim CustomerCase As New CustomerCase
        '    Dim strBody As String
        '    Dim bSuccess As Boolean
        '    Dim CustomerID As Integer

        '    Try
        '        CustomerID = GetCustomerIDByAccountEmail(CustomerEmail, AccountID)

        '        strBody = "Your CharBroil Account password has been reset.  Please click on the following link to create a new password.<br><br>"
        '        strBody = strBody & "<a href=""" & String.Format(ConfigurationManager.AppSettings("ApplicationHome") & "/CustomerHomeWebsite/CustomerHomePasswordUpdate.aspx?source={0}&CustomerID={1}&CustomerIDKey={2}&AccountID={3}&AccountIDKey={4}", "email", CustomerID, EncryptQueryStringValue(CustomerID), AccountID, EncryptQueryStringValue(AccountID)) & """>Set my new password</a>"

        '        'Send email
        '        sEmail.MailToAddress = CustomerEmail
        '        sEmail.MailFromAddress = "CustomerSupport@Charbroil.com"
        '        sEmail.Subject = "Your CharBroil Account"
        '        sEmail.Body = strBody
        '        sEmail.SMTPServer = ConfigurationManager.AppSettings("SMTPServer")
        '        sEmail.Priority = CMS_BL.EmailPriorities.Normal
        '        sEmail.BodyFormat = CMS_BL.EmailBodyFormat.HTML
        '        sEmail.ReplyTo = "CustomerSupport@Charbroil.com"
        '        sEmail.IncludeRedirectReplyLink = False
        '        bSuccess = CustomerCase.SendCaseCorrespondenceEmail(sEmail)
        '    Catch ex As Exception
        '        _strErrorMessage = Err.ToString       'Set the Classes Error Message
        '        _intErrorNumber = 0       'Set the Classes Error Number
        '        bSuccess = False
        '    Finally
        '        CustomerCase = Nothing
        '    End Try

        '    Return bSuccess
        'End Function

        Public Function RegisterProduct(ByRef sProductRegistration As sProductRegistration, Optional ByVal ReserveSerialNumber As Boolean = False) As Boolean

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLTran As SqlTransaction = Nothing
            ' SQLTransaction Object - Used to guarantee that ALL or None of the function Processes  

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim sEmail As sEmail = Nothing

            Dim sCustomer As sCustomer = Nothing
            Dim strBody As String = ""
            '*********************************************************************************************
            ' ****************** To Add a Product Registration Two Steps are Required: ********************
            '   - 1. Create a New Customer Product Ownership Entry 
            '   - 2. Create a New Customer History 
            '   
            '*********************************************************************************************

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database
                SQLTran = SQLConn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin your Transaction


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection
                SQLCommand.Transaction = SQLTran                     'Set the Transaction



                'Step One - Create a Customer Ownership

                'Set the Specific Command Information for Step One
                SQLCommand.CommandText = "LLFBPS..[AddCustomerRegisteredProduct]"  'Stored Procedure Name


                'Preset empty values
                If IsNothing(sProductRegistration.FileName) Then sProductRegistration.FileName = ""
                If IsNothing(sProductRegistration.FilePath) Then sProductRegistration.FilePath = ""
                If IsNothing(sProductRegistration.MimeType) Then sProductRegistration.MimeType = ""
                'Stored Procedure Paramaters - Set their Values using the "sProductRegistration" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SerialNumber", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.SerialNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PurchasePrice", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.PurchasePrice))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateOfPurchase", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.DateOfPurchase))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OwnershipConfirmed", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.OwnershipConfirmed))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Now))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateConfirmed", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.DateConfirmed))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductConfirmationSourceID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.ConfirmationSourceID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.IsActive))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RegistrationSourceID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.RegistrationSourceID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RetailerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.RetailerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Reviewed", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.Reviewed))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.Description))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FilePath", System.Data.SqlDbType.VarChar, 500, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.FilePath))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FileName", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.FileName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MimeType", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.MimeType))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@POPRequestedCaseID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.POPRequestedCaseID))

                If sProductRegistration.POPRequestedDate <> "12:00:00 AM" Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@POPRequestedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.POPRequestedDate))
                Else
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@POPRequestedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, DBNull.Value))
                End If

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@POPRequestedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.POPRequestedBy))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerOwnershipID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.CustomerOwnershipID))


                SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure

                'Pass the new "CustomerOwnershipID" back to the Structure
                sProductRegistration.CustomerOwnershipID = SQLCommand.Parameters("@CustomerOwnershipID").Value


                'If there was a problem with Step One the New "CustomerOwnershipID" will NOT be greater then 0
                If sProductRegistration.CustomerOwnershipID > 0 Then  'IF True Continue to Step Two

                    sProductRegistration._LogInformation &= String.Format(" -- Customer # {0} registered product {1} - (Serial Number -{2})", sProductRegistration.CustomerID, sProductRegistration.ProductModelNumber, sProductRegistration.SerialNumber)





                Else 'Failed Step One - Adding the New Registration


                    _strErrorMessage = "Internal Error adding Product Registration (Not System Error)"  'Set the Classes Error Message
                    _intErrorNumber = 0         'Set the Classes Error Number
                    bSuccess = False           'Set the "Succeed" Boolean to False
                    SQLTran.Rollback()           'RollBack all SQLCommand Inserts


                End If

                '*************************************************************************************************
                'IF BOTH STEPS processed successfully then COMMIT the SQLCOmmand INSERTS
                If bSuccess = True Then SQLTran.Commit()
                '*************************************************************************************************

                'Close Connection
                SQLConn.Close()



                'Reserve the Serial Number if necessary
                If ReserveSerialNumber = True AndAlso bSuccess = True Then
                    Dim sRegisteredSerialNumber As sRegisteredSerialNumber = Nothing

                    With sProductRegistration
                        sRegisteredSerialNumber.CustomerID = .CustomerID
                        sRegisteredSerialNumber.CustomerOwnershipID = .CustomerOwnershipID
                        sRegisteredSerialNumber.DateCreated = Now()
                        sRegisteredSerialNumber.ProductID = .ProductID
                        sRegisteredSerialNumber.ProductModelNumber = .ProductModelNumber
                        sRegisteredSerialNumber.SerialNumber = .SerialNumber.Trim
                        sRegisteredSerialNumber.SerialNumberConfirmed = True

                    End With
                    Me.AddRegisteredSerialNumber(sRegisteredSerialNumber)


                End If


                'Send Registration Email
                'GetCustomer(sCustomer, sProductRegistration.CustomerID)
                'GetRegistration(sProductRegistration, sProductRegistration.CustomerID, sProductRegistration.ProductID)
                'If InStr(sCustomer.Email, "@") Then

                '    strBody = "Thank you.  We have received your registration for the following product:<br><br><font style='size:8pt;'>"
                '    If sProductRegistration.ProductName <> "" Then strBody += "Product:  <b>" & sProductRegistration.ProductName & "</b><br>"
                '    If sProductRegistration.ProductModelNumber <> "" Then strBody += "Product Model Number:  <b>" & sProductRegistration.ProductModelNumber & "</b><br>"
                '    If sProductRegistration.SerialNumber <> "" Then strBody += "Serial Number:  <b>" & sProductRegistration.SerialNumber & "</b><br></font><br><br>"
                '    strBody += "For more information on your product, <a href=""" & String.Format(ConfigurationManager.AppSettings("ApplicationHome") & "/CustomerHomeWebsite/ProductSupport.aspx?ProductID={0}", sProductRegistration.ProductID) & """>Click here</a>" & vbCrLf & " or visit http://www.charbroil.com."

                '    With sEmail
                '        .CustomerID = sProductRegistration.CustomerID
                '        .BodyFormat = EmailBodyFormat.HTML
                '        .MailToAddress = sCustomer.Email
                '        .ReplyTo = "CustomerSupport@Charbroil.com"
                '        .MailFromAddress = "CustomerSupport@Charbroil.com"
                '        .Priority = EmailPriorities.Normal
                '        .Subject = "Your Product Registration"
                '        .Body = strBody
                '        .SMTPServer = ConfigurationManager.AppSettings("SMTPServer").ToString
                '    End With

                '    SendRegistrationConfirmationEmail(sEmail)
                'End If




                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts


                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing


            End Try

            'Set returning Boolean
            Return bSuccess


        End Function


        Public Function UpdateRegisterProduct(ByRef sProductRegistration As sProductRegistration, Optional ByVal ReserveSerialNumber As Boolean = False) As Boolean

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLTran As SqlTransaction = Nothing
            ' SQLTransaction Object - Used to guarantee that ALL or None of the function Processes  

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Update a Product Registration Two Steps are Required: ********************
            '   - 1. Update the Customer Product Ownership Entry 
            '   - 2. Create a New Customer History 
            '   
            '*********************************************************************************************

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database
                SQLTran = SQLConn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin your Transaction


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection
                SQLCommand.Transaction = SQLTran                     'Set the Transaction



                'Step One - Create a Customer Ownership

                'Set the Specific Command Information for Step One
                SQLCommand.CommandText = "LLFBPS..[UpdateCustomerProductOwnership]"  'Stored Procedure Name


                If IsNothing(sProductRegistration.FileName) Then sProductRegistration.FileName = ""
                If IsNothing(sProductRegistration.FilePath) Then sProductRegistration.FilePath = ""
                If IsNothing(sProductRegistration.MimeType) Then sProductRegistration.MimeType = ""

                'Stored Procedure Paramaters - Set their Values using the "sProductRegistration" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerOwnershipID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.CustomerOwnershipID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SerialNumber", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.SerialNumber))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PurchasePrice", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.PurchasePrice))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateOfPurchase", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.DateOfPurchase))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OwnershipConfirmed", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.OwnershipConfirmed))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Now))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateConfirmed", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.DateConfirmed))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductConfirmationSourceID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.ConfirmationSourceID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.IsActive))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RegistrationSourceID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.RegistrationSourceID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RetailerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.RetailerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Reviewed", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.Reviewed))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.Description))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FilePath", System.Data.SqlDbType.VarChar, 500, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.FilePath))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@FileName", System.Data.SqlDbType.VarChar, 75, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.FileName))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@MimeType", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.MimeType))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@POPRequestedCaseID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.POPRequestedCaseID))
                If sProductRegistration.POPRequestedDate <> "12:00:00 AM" Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@POPRequestedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.POPRequestedDate))
                Else
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@POPRequestedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, DBNull.Value))
                End If
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@POPRequestedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, True, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sProductRegistration.POPRequestedBy))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerOwnershipUpdatedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure


                'If there was a problem with Step One the "CustomerOwnershipUpdatedDate" will NOT be a date
                If IsDate(SQLCommand.Parameters("@CustomerOwnershipUpdatedDate").Value) Then 'IF True Continue to Step Two


                    'Set the Specific Command Information for Step Two
                    SQLCommand.Parameters.Clear() 'Clear the Parameters 
                    SQLCommand.CommandText = "LLFBPS..[AddCustomerHistory]" 'Stored Procedure Name

                    sProductRegistration._LogInformation &= String.Format(" -- Customer # {0} updated registration product {1} - (Serial Number -{2})", sProductRegistration.CustomerID, sProductRegistration.ProductModelNumber, sProductRegistration.SerialNumber)

                    'Stored Procedure Paramaters - Set their Values using the Structure and/or Private Variables ( m_lngCustomerID)
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sProductRegistration.CustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, String.Format("Customer # {0} registered product {1} - (Serial Number -{2})", sProductRegistration.CustomerID, sProductRegistration.ProductModelNumber, sProductRegistration.SerialNumber)))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, 0))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Now))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerHistoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                    SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure



                    'If there was a problem with Step Two the New "CustomerHistoryID" will NOT be greater then 0

                    If Not SQLCommand.Parameters("@CustomerHistoryID").Value > 0 Then 'If True then if FAILED Step Three

                        _strErrorMessage = "Error adding Customer History"             'Set the Classes Error Message
                        _intErrorNumber = 0                                        'Set the Classes Error Number
                        bSuccess = False                                          'Set the "Succeed" Boolean to False
                        SQLTran.Rollback()                                          'RollBack all SQLCommand Inserts

                    End If



                Else 'Failed Step One - Adding the New Registration


                    _strErrorMessage = "Internal Error updating Product Registration (Not System Error)"  'Set the Classes Error Message
                    _intErrorNumber = 0         'Set the Classes Error Number
                    bSuccess = False           'Set the "Succeed" Boolean to False
                    SQLTran.Rollback()           'RollBack all SQLCommand Inserts


                End If

                '*************************************************************************************************
                'IF BOTH STEPS processed successfully then COMMIT the SQLCOmmand INSERTS
                If bSuccess = True Then SQLTran.Commit()
                '*************************************************************************************************

                'Close Connection
                SQLConn.Close()



                'Reserve the Serial Number if necessary
                If ReserveSerialNumber = True AndAlso bSuccess = True Then
                    Dim sRegisteredSerialNumber As sRegisteredSerialNumber = Nothing

                    With sProductRegistration
                        sRegisteredSerialNumber.CustomerID = .CustomerID
                        sRegisteredSerialNumber.CustomerOwnershipID = .CustomerOwnershipID
                        sRegisteredSerialNumber.DateCreated = Now()
                        sRegisteredSerialNumber.ProductID = .ProductID
                        sRegisteredSerialNumber.ProductModelNumber = .ProductModelNumber
                        sRegisteredSerialNumber.SerialNumber = .SerialNumber.Trim
                        sRegisteredSerialNumber.SerialNumberConfirmed = True

                    End With
                    Me.AddRegisteredSerialNumber(sRegisteredSerialNumber)


                End If


                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts


                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function




        Public Function AddAccount(ByRef sAccount As sAccount) As Boolean

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLTran As SqlTransaction = Nothing
            ' SQLTransaction Object - Used to guarantee that ALL or None of the function Processes  

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add an Account Two Steps are Required: ********************
            '   - 1. Create a New Customer Account JUNC Entry 
            '   - 2. Create a New Customer History 
            '   
            '*********************************************************************************************

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database
                SQLTran = SQLConn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin your Transaction


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection
                SQLCommand.Transaction = SQLTran                     'Set the Transaction



                'Step One - Create a Customer Ownership

                'Set the Specific Command Information for Step One
                SQLCommand.CommandText = "LLFBPS..[AddCustomerAccount]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AccountID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.AccountID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Username", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.Username))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Password", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.Password))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Now))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.IsActive))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerAccountID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.CustomerAccountID))


                SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure

                'Pass the new "CustomerAccountID" back to the Structure
                sAccount.CustomerAccountID = SQLCommand.Parameters("@CustomerAccountID").Value


                'If there was a problem with Step One the New "CustomerAccountID" will NOT be greater then 0
                If sAccount.CustomerAccountID > 0 Then   'IF True Continue to Step Two


                    'Set the Specific Command Information for Step Two
                    SQLCommand.Parameters.Clear() 'Clear the Parameters 
                    SQLCommand.CommandText = "LLFBPS..[AddCustomerHistory]" 'Stored Procedure Name

                    'Stored Procedure Paramaters - Set their Values using the Structure and/or Private Variables ( m_lngCustomerID)
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sAccount.CustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, String.Format("Customer # {0} joined Account {1})", sAccount.CustomerID, sAccount.Name)))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, 0))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Now))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerHistoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                    SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure



                    'If there was a problem with Step Two the New "CustomerHistoryID" will NOT be greater then 0

                    If Not SQLCommand.Parameters("@CustomerHistoryID").Value > 0 Then 'If True then if FAILED Step Three

                        _strErrorMessage = "Error adding Customer History"             'Set the Classes Error Message
                        _intErrorNumber = 0                                        'Set the Classes Error Number
                        bSuccess = False                                          'Set the "Succeed" Boolean to False
                        SQLTran.Rollback()                                          'RollBack all SQLCommand Inserts

                    End If



                Else 'Failed Step One - Adding the New Customer Account 


                    _strErrorMessage = "Internal Error adding Customer Account (Not System Error)"  'Set the Classes Error Message
                    _intErrorNumber = 0         'Set the Classes Error Number
                    bSuccess = False           'Set the "Succeed" Boolean to False
                    SQLTran.Rollback()           'RollBack all SQLCommand Inserts


                End If

                '*************************************************************************************************
                'IF BOTH STEPS processed successfully then COMMIT the SQLCOmmand INSERTS
                If bSuccess = True Then SQLTran.Commit()
                '*************************************************************************************************

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts


                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function


        Public Function AddNote(ByRef sCustomerNote As sCustomerNote) As Boolean

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLTran As SqlTransaction = Nothing
            ' SQLTransaction Object - Used to guarantee that ALL or None of the function Processes  

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add an Note Two Steps are Required: ********************
            '   - 1. Create a New Customer Note
            '   - 2. Create a New Customer History 
            '   
            '*********************************************************************************************

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database
                SQLTran = SQLConn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin your Transaction


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection
                SQLCommand.Transaction = SQLTran                     'Set the Transaction



                'Step One - Create a Customer Ownership

                'Set the Specific Command Information for Step One
                SQLCommand.CommandText = "LLFBPS..[AddCustomerNote]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerNoteTypeID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerNoteTypeID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CompanyID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.BrandID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.Name))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 5000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.Description))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CreatedBy))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.IsActive))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerNoteID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))



                SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure

                'Pass the new "CustomerNoteID" back to the Structure
                sCustomerNote.CustomerNoteID = SQLCommand.Parameters("@CustomerNoteID").Value


                'If there was a problem with Step One the New "CustomerNoteID" will NOT be greater then 0
                If sCustomerNote.CustomerNoteID > 0 Then    'IF True Continue to Step Two


                    'Set the Specific Command Information for Step Two
                    SQLCommand.Parameters.Clear() 'Clear the Parameters 
                    SQLCommand.CommandText = "LLFBPS..[AddCustomerHistory]" 'Stored Procedure Name

                    'Stored Procedure Paramaters - Set their Values using the Structure and/or Private Variables ( m_lngCustomerID)
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomerNote.CustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, String.Format("Customer # {0} had a note created titled - {1})", sCustomerNote.CustomerID, sCustomerNote.Name)))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomerNote.CreatedBy))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Now))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerHistoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                    SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure



                    'If there was a problem with Step Two the New "CustomerHistoryID" will NOT be greater then 0

                    If Not SQLCommand.Parameters("@CustomerHistoryID").Value > 0 Then 'If True then if FAILED Step Three

                        _strErrorMessage = "Error adding Customer History"             'Set the Classes Error Message
                        _intErrorNumber = 0                                        'Set the Classes Error Number
                        bSuccess = False                                          'Set the "Succeed" Boolean to False
                        SQLTran.Rollback()                                          'RollBack all SQLCommand Inserts

                    End If



                Else 'Failed Step One - Adding the New Customer Note


                    _strErrorMessage = "Internal Error adding Customer Account (Not System Error)"  'Set the Classes Error Message
                    _intErrorNumber = 0         'Set the Classes Error Number
                    bSuccess = False           'Set the "Succeed" Boolean to False
                    SQLTran.Rollback()           'RollBack all SQLCommand Inserts


                End If

                '*************************************************************************************************
                'IF BOTH STEPS processed successfully then COMMIT the SQLCOmmand INSERTS
                If bSuccess = True Then SQLTran.Commit()
                '*************************************************************************************************

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts


                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function


        Public Function UpdateNote(ByRef sCustomerNote As sCustomerNote) As Boolean

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLTran As SqlTransaction = Nothing
            ' SQLTransaction Object - Used to guarantee that ALL or None of the function Processes  

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Update an Note Two Steps are Required: ********************
            '   - 1. Update Customer Note
            '   - 2. Create a New Customer History 
            '   
            '*********************************************************************************************

            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database
                SQLTran = SQLConn.BeginTransaction(IsolationLevel.ReadCommitted) 'Begin your Transaction


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection
                SQLCommand.Transaction = SQLTran                     'Set the Transaction



                'Step One - Create a Customer Note

                'Set the Specific Command Information for Step One
                SQLCommand.CommandText = "LLFBPS..[UpdateCustomerNote]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerNoteID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerNoteID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerNoteTypeID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerNoteTypeID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CompanyID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.BrandID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.ProductID))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Name", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.Name))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 5000, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.Description))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CreatedBy))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.IsActive))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerNoteUpdatedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))



                SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure


                'If there was a problem with Step One the New "CustomerNoteUpdatedDate" will NOT be a date
                If IsDate(SQLCommand.Parameters("@CustomerNoteUpdatedDate").Value) Then    'IF True Continue to Step Two


                    'Set the Specific Command Information for Step Two
                    SQLCommand.Parameters.Clear() 'Clear the Parameters 
                    SQLCommand.CommandText = "LLFBPS..[AddCustomerHistory]" 'Stored Procedure Name

                    'Stored Procedure Paramaters - Set their Values using the Structure and/or Private Variables ( m_lngCustomerID)
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomerNote.CustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, String.Format("Customer # {0} had a noteID {2} updated titled - {1})", sCustomerNote.CustomerID, sCustomerNote.Name, sCustomerNote.CustomerNoteID)))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, sCustomerNote.CreatedBy))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateCreated", System.Data.SqlDbType.DateTime, 8, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, Now))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerHistoryID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                    SQLCommand.ExecuteNonQuery() 'Execute the Stored Procedure



                    'If there was a problem with Step Two the New "CustomerHistoryID" will NOT be greater then 0

                    If Not SQLCommand.Parameters("@CustomerHistoryID").Value > 0 Then 'If True then if FAILED Step Three

                        _strErrorMessage = "Error adding Customer History"             'Set the Classes Error Message
                        _intErrorNumber = 0                                        'Set the Classes Error Number
                        bSuccess = False                                          'Set the "Succeed" Boolean to False
                        SQLTran.Rollback()                                          'RollBack all SQLCommand Inserts

                    End If



                Else 'Failed Step One - Adding the New Customer Note


                    _strErrorMessage = "Internal Error adding Customer Account (Not System Error)"  'Set the Classes Error Message
                    _intErrorNumber = 0         'Set the Classes Error Number
                    bSuccess = False           'Set the "Succeed" Boolean to False
                    SQLTran.Rollback()           'RollBack all SQLCommand Inserts


                End If

                '*************************************************************************************************
                'IF BOTH STEPS processed successfully then COMMIT the SQLCOmmand INSERTS
                If bSuccess = True Then SQLTran.Commit()
                '*************************************************************************************************

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts


                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                SQLTran.Rollback()                                    'RollBack all SQLCommand Inserts

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function




        Public Function GetAccount(ByRef sAccount As sAccount, ByVal AuthenticationMode As CustomerAccountAuthenticationMode) As Boolean





            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = False  'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            Dim SQLDataReader As SqlDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[GetCustomerAccount]"  'Stored Procedure Name


                'Stored Procedure Paramaters - Set their Values using the Structure and Authentication Mode

                Select Case AuthenticationMode

                    Case CustomerAccountAuthenticationMode.UsernameAndPassword

                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Username", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.Username))
                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Password", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.Password))

                    Case CustomerAccountAuthenticationMode.CustomerAccountID
                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerAccountID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.CustomerAccountID))


                    Case CustomerAccountAuthenticationMode.CustomerAndAccountID
                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.CustomerID))
                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AccountID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.AccountID))

                    Case CustomerAccountAuthenticationMode.UsernameAndAccountID
                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Username", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.Username))
                        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AccountID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.AccountID))


                End Select




                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


                If SQLDataReader.Read = True Then
                    With sAccount
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("FirstName")) Then .CustomerFirstname = SQLDataReader("FirstName")
                        .CustomerID = SQLDataReader("CustomerID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("LastName")) Then .CustomerLastname = SQLDataReader("LastName")
                        If IsDate(.DateCreated) Then .DateCreated = SQLDataReader("DateCreated")
                        .IsActive = SQLDataReader("Active")

                        .IsBusiness = SQLDataReader("IsBusiness")
                        .CompanyName = SQLDataReader("CompanyName").ToString
                        Try
                            .AccountStatusID = SQLDataReader("AccountStatusID").ToString
                        Catch ex As Exception
                            .AccountStatusID = AccountStatus.Unknown
                        End Try

                        Try
                            .PurchaseStatusID = SQLDataReader("PurchaseStatusID").ToString
                        Catch ex As Exception
                            .PurchaseStatusID = PurchaseStatus.NoInvoice_Web
                        End Try
                        .PasswordExpirationDate = SQLDataReader("PasswordExpirationDate")

                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("FirstName")) Then .Name = SQLDataReader("FirstName") & " " & SQLDataReader("LastName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Password")) Then .Password = SQLDataReader("Password")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Email")) Then .Username = SQLDataReader("Email")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("PriceListID")) Then .PriceListID = SQLDataReader("PriceListID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Zip")) Then .CustomerZip = SQLDataReader("Zip")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("LastLoginDate")) Then .LastLoginDate = SQLDataReader("LastLoginDate") Else .LastLoginDate = Now
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("LoginDate")) Then .LoginDate = SQLDataReader("LoginDate") Else .LoginDate = Now




                    End With

                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function



        'Public Function GetNote(ByRef sCustomerNote As sCustomerNote) As Boolean





        '    'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    Dim bSuccess As Boolean = False  'Used to confirm that the Function Succesfully Processed

        '    'Set Connection Object
        '    Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

        '    Dim SQLCommand As New SqlCommand 'SQLCommand Object
        '    Dim SQLDataReader As SqlDataReader
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '    Try 'Only one "Try" statement 

        '        SQLConn.Open() 'Open Database


        '        'Set the Basic Command Information (This Command object is used Two times)
        '        SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
        '        SQLCommand.Connection = SQLConn                      'Set the Connection


        '        'Set the Specific Command Information 
        '        SQLCommand.CommandText = "LLFBPS..[GetCustomerNotes]"  'Stored Procedure Name


        '        If sCustomerNote.CustomerID > 0 Then SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerID))
        '        If sCustomerNote.CustomerNoteID > 0 Then SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerNoteID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerNoteID))



        '        'Set the DataReader
        '        SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


        '        If SQLDataReader.Read = True Then
        '            With sCustomerNote

        '           .
        '                .BrandID = SQLDataReader("BrandID")
        '                .BrandName = SQLDataReader("BrandName")
        '                .CompanyID = SQLDataReader("CompanyID")
        '                .CompanyName = SQLDataReader("CompanyName")
        '                .CustomerNoteID = SQLDataReader("CustomerNoteID")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("CustomerFirstName")) Then .CustomerFirstname = SQLDataReader("CustomerFirstName")
        '                .CustomerID = SQLDataReader("CustomerID")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("CustomerLastName")) Then .CustomerLastname = SQLDataReader("CustomerLastName")
        '                If IsDate(.DateCreated) Then .DateCreated = SQLDataReader("DateCreated")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
        '                .IsActive = SQLDataReader("Active")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Password")) Then .Password = SQLDataReader("Password")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("UserName")) Then .Username = SQLDataReader("UserName")


        '            End With

        '            bSuccess = True
        '        End If

        '        'Close Connection
        '        SQLConn.Close()

        '        'SQL ERROR CATCH
        '    Catch SQLErr As SqlException
        '        bSuccess = False                                    'Set the "Succeed" Boolean to False
        '        _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
        '        _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

        '        'MISC ERROR CATCH
        '    Catch Err As Exception
        '        bSuccess = False                                    'Set the "Succeed" Boolean to False

        '        _strErrorMessage = Err.ToString                      'Set the Classes Error Message
        '        _intErrorNumber = 0                                  'Set the Classes Error Number


        '    Finally
        '        'Confirm that The SQLDB Connection is closed
        '        If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
        '        SQLCommand = Nothing
        '        SQLConn = Nothing

        '    End Try

        '    'Set returning Boolean
        '    Return bSuccess

        'End Function



        'Public Function GetAccountByUserName(ByRef sAccount As sAccount) As Boolean


        '    'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    Dim bSuccess As Boolean = False  'Used to confirm that the Function Succesfully Processed

        '    'Set Connection Object
        '    Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

        '    Dim SQLCommand As New SqlCommand 'SQLCommand Object
        '    Dim SQLDataReader As SqlDataReader
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '    Try 'Only one "Try" statement 

        '        SQLConn.Open() 'Open Database


        '        'Set the Basic Command Information (This Command object is used Two times)
        '        SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
        '        SQLCommand.Connection = SQLConn                      'Set the Connection


        '        'Set the Specific Command Information 
        '        SQLCommand.CommandText = "LLFBPS..[GetCustomerAccount]"  'Stored Procedure Name


        '        'Stored Procedure Paramaters - Set their Values using the Structure and Authentication Mode


        '        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Username", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.Username))
        '        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AccountID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sAccount.AccountID))






        '        'Set the DataReader
        '        SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


        '        If SQLDataReader.Read = True Then
        '            With sAccount

        '                .AccountID = SQLDataReader("AccountID")
        '                .BrandID = SQLDataReader("BrandID")
        '                .BrandName = SQLDataReader("BrandName")
        '                .CompanyID = SQLDataReader("CompanyID")
        '                .CompanyName = SQLDataReader("CompanyName")
        '                .CustomerAccountID = SQLDataReader("CustomerAccountID")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("CustomerFirstName")) Then .CustomerFirstname = SQLDataReader("CustomerFirstName")
        '                .CustomerID = SQLDataReader("CustomerID")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("CustomerLastName")) Then .CustomerLastname = SQLDataReader("CustomerLastName")
        '                If IsDate(.DateCreated) Then .DateCreated = SQLDataReader("DateCreated")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
        '                .IsActive = SQLDataReader("Active")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Password")) Then .Password = SQLDataReader("Password")
        '                If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("UserName")) Then .Username = SQLDataReader("UserName")


        '            End With

        '            bSuccess = True
        '        End If

        '        'Close Connection
        '        SQLConn.Close()

        '        'SQL ERROR CATCH
        '    Catch SQLErr As SqlException
        '        bSuccess = False                                    'Set the "Succeed" Boolean to False
        '        _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
        '        _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

        '        'MISC ERROR CATCH
        '    Catch Err As Exception
        '        bSuccess = False                                    'Set the "Succeed" Boolean to False

        '        _strErrorMessage = Err.ToString                      'Set the Classes Error Message
        '        _intErrorNumber = 0                                  'Set the Classes Error Number


        '    Finally
        '        'Confirm that The SQLDB Connection is closed
        '        If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
        '        SQLCommand = Nothing
        '        SQLConn = Nothing

        '    End Try

        '    'Set returning Boolean
        '    Return bSuccess

        'End Function
        Public Function GetRegistration(ByRef sProductRegistration As sProductRegistration, ByVal CustomerID As Long, ByVal ProductID As Long) As Boolean
            _lngCustomerID = CustomerID
            _lngProductID = ProductID

            Return _GetRegistration(sProductRegistration, FindProductRegistrationType.ByCustomerAndProduct)

        End Function

        Public Function GetRegistration(ByRef sProductRegistration As sProductRegistration, ByVal CustomerOwnerShipID As Long) As Boolean
            _lngCustomerOwnershipID = CustomerOwnerShipID
            Return _GetRegistration(sProductRegistration, FindProductRegistrationType.ByRegistrationID)

        End Function

        Private Function _GetRegistration(ByRef sProductRegistration As sProductRegistration, ByVal FindProductRegistrationType As FindProductRegistrationType) As Boolean



            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = False  'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            Dim SQLDataReader As SqlDataReader

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[GetProductRegistration]"  'Stored Procedure Name

                'Stored Procedure Paramaters - 
                If FindProductRegistrationType = FindProductRegistrationType.ByRegistrationID Then
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerOwnerShipID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, _lngCustomerOwnershipID))
                Else
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, _lngCustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, _lngProductID))

                End If

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


                If SQLDataReader.Read = True Then
                    With sProductRegistration


                        'Get the Registration Information


                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        .CustomerID = SQLDataReader("CustomerID")
                        .CustomerOwnershipID = SQLDataReader("CustomerOwnerShipID")
                        If IsDate(SQLDataReader("DateConfirmed")) Then .DateConfirmed = SQLDataReader("DateConfirmed")
                        If IsDate(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If IsDate(SQLDataReader("DateOfPurchase")) Then .DateOfPurchase = SQLDataReader("DateOfPurchase")
                        .ProductID = SQLDataReader("ProductID")
                        .ProductName = SQLDataReader("ProductName")
                        .ProductModelNumber = SQLDataReader("ProductModelNumber")
                        .IsActive = SQLDataReader("Active")
                        .OwnershipConfirmed = SQLDataReader("OwnershipConfirmed")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("ProductConfirmationSourceID")) Then .ConfirmationSourceID = SQLDataReader("ProductConfirmationSourceID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("PurchasePrice")) Then .PurchasePrice = SQLDataReader("PurchasePrice")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("ProductRegistrationSourceID")) Then .RegistrationSourceID = SQLDataReader("ProductRegistrationSourceID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("RetailerID")) Then .RetailerID = SQLDataReader("RetailerID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("RetailerName")) Then .RetailerName = SQLDataReader("RetailerName")
                        .Reviewed = SQLDataReader("Reviewed")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("SerialNumber")) Then .SerialNumber = SQLDataReader("SerialNumber")


                        'Get the Customer Information
                        .Customer.CustomerID = SQLDataReader("CustomerID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("FirstName")) Then .Customer.FirstName = SQLDataReader("FirstName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("MiddleName")) Then .Customer.MiddleName = SQLDataReader("MiddleName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("LastName")) Then .Customer.LastName = SQLDataReader("LastName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Address1")) Then .Customer.Address1 = SQLDataReader("Address1")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Address2")) Then .Customer.Address2 = SQLDataReader("Address2")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("City")) Then .Customer.City = SQLDataReader("City")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("State")) Then .Customer.State = SQLDataReader("State")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Zip")) Then .Customer.Zip = SQLDataReader("Zip")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Country")) Then .Customer.Country = SQLDataReader("Country")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Email")) Then .Customer.Email = SQLDataReader("Email")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("PrimaryPhone")) Then .Customer.PrimaryPhone = SQLDataReader("PrimaryPhone")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("SecondaryPhone")) Then .Customer.SecondaryPhone = SQLDataReader("SecondaryPhone")
                        '.Customer.FullName = String.Format("{0} {1}", .Customer.FirstName, .Customer.LastName)
                        'Format the full name with the correct case
                        .Customer.FullName = String.Format("{0} {1}", UCase(Left(.Customer.FirstName, 1)) + LCase(Mid(.Customer.FirstName, 2)), UCase(Left(.Customer.LastName, 1)) + LCase(Mid(.Customer.LastName, 2)))
                        .Customer.ContactSheet = String.Format("{0} - Addr: {5}, {6}, {7}, {8} - Email: {1} Home#: {2} Work#: {3} Mobile#: {4}", .Customer.FullName, .Customer.Email, .Customer.HomePhone, .Customer.WorkPhone, .Customer.MobilePhone, .Customer.Address1, .Customer.City, .Customer.State, .Customer.Zip)


                        If Not IsDBNull(SQLDataReader("POPRequestedCaseID")) Then .POPRequestedCaseID = SQLDataReader("POPRequestedCaseID")
                        If Not IsDBNull(SQLDataReader("POPRequestedDate")) Then .POPRequestedDate = SQLDataReader("POPRequestedDate")
                        If Not IsDBNull(SQLDataReader("POPRequestedBy")) Then .POPRequestedBy = SQLDataReader("POPRequestedBy")
                        If Not IsDBNull(SQLDataReader("POPRequestedByName")) Then .POPRequestedByName = SQLDataReader("POPRequestedByName")
                        If Not IsDBNull(SQLDataReader("FilePath")) Then .FilePath = SQLDataReader("FilePath")
                        If Not IsDBNull(SQLDataReader("FileName")) Then .FileName = SQLDataReader("FileName")
                        If Not IsDBNull(SQLDataReader("MimeType")) Then .MimeType = SQLDataReader("MimeType")







                    End With

                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function



        Private Function _GetCustomer(ByRef sCustomer As sCustomer) As Boolean





            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = False  'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            Dim SQLDataReader As SqlDataReader

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement 

                SQLConn.Open() 'Open Database


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[GetCustomer]"  'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomer" Structure
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, _lngCustomerID))

                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


                If SQLDataReader.Read = True Then
                    With sCustomer

                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("CompanyName")) Then .CompanyName = SQLDataReader("CompanyName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("FirstName")) Then .FirstName = SQLDataReader("FirstName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("MiddleName")) Then .MiddleName = SQLDataReader("MiddleName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("LastName")) Then .LastName = SQLDataReader("LastName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Gender")) Then .Gender = SQLDataReader("Gender")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Address1")) Then .Address1 = SQLDataReader("Address1")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Address2")) Then .Address2 = SQLDataReader("Address2")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("City")) Then .City = SQLDataReader("City")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("State")) Then .State = SQLDataReader("State")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Zip")) Then .Zip = SQLDataReader("Zip")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Country")) Then .Country = SQLDataReader("Country")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Email")) Then .Email = SQLDataReader("Email")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("PrimaryPhone")) Then .PrimaryPhone = SQLDataReader("PrimaryPhone")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("SecondaryPhone")) Then .SecondaryPhone = SQLDataReader("SecondaryPhone")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("AlternateCustomerNumber")) Then .AlternateCustomerNumber = SQLDataReader("AlternateCustomerNumber")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("AlternateCustomerNumber2")) Then .AlternateCustomerNumber2 = SQLDataReader("AlternateCustomerNumber2")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("AlternateCustomerNumber3")) Then .AlternateCustomerNumber3 = SQLDataReader("AlternateCustomerNumber3")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("AlternateCustomerNumber4")) Then .AlternateCustomerNumber4 = SQLDataReader("AlternateCustomerNumber4")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("UserField1")) Then .UserField1 = SQLDataReader("UserField1")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("UserField2")) Then .UserField2 = SQLDataReader("UserField2")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("UserField3")) Then .UserField3 = SQLDataReader("UserField3")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("IsBusiness")) Then .IsBusiness = SQLDataReader("IsBusiness")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("TaxExemptNumber")) Then .TaxExemptNumber = SQLDataReader("TaxExemptNumber")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("RetailerID")) Then .RetailerID = SQLDataReader("RetailerID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("TaxStatusID")) Then .TaxStatusID = SQLDataReader("TaxStatusID")

                        .CreatedBy = SQLDataReader("CreatedBy")
                        If IsDate(.DateCreated) Then .DateCreated = SQLDataReader("DateCreated")
                        .Active = SQLDataReader("Active")
                        .CustomerID = SQLDataReader("CustomerID")
                        'If Not Microsoft.VisualBasic.IsDBNull(

                        'Format the full name with the correct case
                        .FullName = String.Format("{0} {1}", UCase(Left(.FirstName, 1)) + LCase(Mid(.FirstName, 2)), UCase(Left(.LastName, 1)) + LCase(Mid(.LastName, 2)))
                        .ContactSheet = String.Format("{0} - Addr: {4}, {5}, {6}, {7} - Email: {1} Primary Phone#: {2} Secondary Phone#: {3} ", .FullName, .Email, .PrimaryPhone, .SecondaryPhone, .Address1, .City, .State, .Zip)




                    End With

                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

        End Function




        Public Function AddRegisteredSerialNumber(ByRef sRegisteredSerialNumber As sRegisteredSerialNumber) As Boolean

            'Function: AddRegisteredSerialNumber ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddRegisteredSerialNumber
            '       Parent Class:  Product
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sRegisteredSerialNumber Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  03/16/05
            '       Description: Adds a Registered Serial Number
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add  One Step is Required: ********************
            '   - 1. Add the Registered Serial Number
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[AddRegisteredSerialNumber]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure

                With sRegisteredSerialNumber
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CustomerID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SerialNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .SerialNumber))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductModelNumber", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .ProductModelNumber))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CompanyID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .BrandID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SerialNumberConfirmed", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .SerialNumberConfirmed))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerOwnershipID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, .CustomerOwnershipID))
                    SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RegisteredSerialNumberID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))



                End With

                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "ID" back to the Structure
                sRegisteredSerialNumber.RegisteredSerialNumberID = SQLCommand.Parameters("@RegisteredSerialNumberID").Value

                If Not SQLCommand.Parameters("@RegisteredSerialNumberID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Registered Serial (Not System Error)"     'Set the Classes Error Message
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









        Public Function GetNotes(ByRef sCustomerNote As sCustomerNote) As Boolean



            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = False  'Used to confirm that the Function Succesfully Processed

            'Set Connection Object
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand 'SQLCommand Object
            Dim SQLDataReader As SqlDataReader
            Dim sCustomer As sCustomer = Nothing 'Structure to get the customer information

            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Only one "Try" statement

                SQLConn.Open() 'Open Database


                'Set the Basic Command Information (This Command object is used Two times)
                SQLCommand.CommandType = CommandType.StoredProcedure 'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                      'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[GetCustomerNotes]"  'Stored Procedure Name


                If sCustomerNote.CustomerID > 0 Then SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.VarChar, 150, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerID))
                If sCustomerNote.CustomerNoteID > 0 Then SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerNoteID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, sCustomerNote.CustomerNoteID))




                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.Default)


                If SQLDataReader.Read = True Then
                    With sCustomerNote


                        'Get the Note Information
                        .CustomerNoteID = SQLDataReader("CustomerNoteID")
                        .CustomerNoteTypeID = SQLDataReader("CustomerNoteTypeID")
                        .CustomerNoteTypeName = SQLDataReader("CustomerNoteTypeName")
                        .CustomerID = SQLDataReader("CustomerID")
                        .CreatedBy = SQLDataReader("CreatedBy")
                        .CreatedByName = SQLDataReader("CreatedByName")
                        .DateCreated = SQLDataReader("DateCreated")
                        .IsActive = SQLDataReader("Active")

                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("CompanyID")) Then .CompanyID = SQLDataReader("CompanyID")
                        If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("CompanyName")) Then .CompanyName = SQLDataReader("CompanyName")


                        'Check to see if a product was created
                        If IsNumeric(SQLDataReader("ProductID")) AndAlso SQLDataReader("ProductID") > 0 Then
                            .ProductID = SQLDataReader("ProductID")
                            If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("ProductName")) Then .ProductName = SQLDataReader("ProductName")
                            If Not Microsoft.VisualBasic.IsDBNull(SQLDataReader("ProductModelNumber")) Then .ProductModelNumber = SQLDataReader("ProductModelNumber")
                        End If



                    End With

                    bSuccess = True
                End If

                'Close Connection
                SQLConn.Close()


                If bSuccess = True Then
                    'Get the customer Information
                    If Me.GetCustomer(sCustomer, sCustomerNote.CustomerID) = True Then
                        sCustomerNote.Customer = sCustomer
                    End If
                End If

                'SQL ERROR CATCH
            Catch SQLErr As SqlException
                bSuccess = False                                    'Set the "Succeed" Boolean to False
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

                'MISC ERROR CATCH
            Catch Err As Exception
                bSuccess = False                                    'Set the "Succeed" Boolean to False

                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number


            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning Boolean
            Return bSuccess

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
        Public Property ProductCategoryID() As Long
            Get
                Return _lngProductCategoryID  'Gets Local _lngProductCategoryID Variable
            End Get
            Set(ByVal Value As Long)
                _lngProductCategoryID = Value 'Sets Local _lngProductCategoryID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
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

            'CreatedProductCategory = AddProductCategory(sProductCategory)

        End Sub




#End Region

#Region "Functions & Procedures"

        'Public Function AddProductCategory()

        'End Function

        'Public Function GetProductCategory(ByRef sProductCategory As sProductCategory) As Boolean

        '    'Function: GetProductCategory ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    '
        '    '       Name: GetProductCategory
        '    '       Parent Class:  ProductCategory
        '    '       Parent File: WorkElements.vb
        '    '       Parent Project: BPS
        '    '       Parent Solution: BPS
        '    '
        '    '       Return Value: BOOLEAN
        '    '       Required Parameters(1): sProductCategory Structure(ByRef) 
        '    '       Optional Parameters(0):
        '    '
        '    '       Created By:  Vincent Clover
        '    '       Created on:  11/18/05
        '    '       Description: Gets a Product Category
        '    '
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        '    If sProductCategory.ProductCategoryID > 0 Then
        '        _lngProductCategoryID = sProductCategory.ProductCategoryID
        '    End If

        '    'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        '    Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

        '    Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

        '    Dim SQLCommand As New SqlCommand   'SQLCommand Object
        '    Dim SQLDataReader As SqlDataReader 'SQLDataReader
        '    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


        '    Try 'Only one "Try" statement 

        '        SQLConn.Open() 'Open Database

        '        'Set the Basic Command Information 
        '        SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
        '        SQLCommand.Connection = SQLConn                          'Set the Connection
        '        SQLCommand.CommandText = "LLFBPS..[GetProductSeriesByID]"              'Stored Procedure Name
        '        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductSeriesID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngProductSeriesID))

        '        'Set the DataReader
        '        SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


        '        'Populate the Product Category Structure
        '        SQLDataReader.Read()
        '        If SQLDataReader.HasRows Then

        '            With sProductCategory
        '                'Fill  Information

        '                If Not IsDBNull(SQLDataReader("ProductSeriesID")) Then .ProductSeriesID = SQLDataReader("ProductSeriesID")
        '                If Not IsDBNull(SQLDataReader("Name")) Then .Name = SQLDataReader("Name")
        '                If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
        '                If Not IsDBNull(SQLDataReader("Freight")) Then .Freight = SQLDataReader("Freight")
        '                If Not IsDBNull(SQLDataReader("IsSupported")) Then .IsSupported = SQLDataReader("IsSupported")
        '                If Not IsDBNull(SQLDataReader("Specifications")) Then .Specifications = SQLDataReader("Specifications")
        '                If Not IsDBNull(SQLDataReader("PrimaryResourceImageID")) Then .PrimaryResourceImageID = SQLDataReader("PrimaryResourceImageID")
        '                'If Not IsDBNull(SQLDataReader("PrimaryResourceImageName")) Then .PrimaryResourceImageName = SQLDataReader("PrimaryResourceImageName")
        '                If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
        '                If Not IsDBNull(SQLDataReader("CreatedName")) Then .CreatedByName = SQLDataReader("CreatedName")
        '                If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
        '                If Not IsDBNull(SQLDataReader("Active")) Then .Active = SQLDataReader("Active")

        '                If Not IsDBNull(SQLDataReader("CompanyID")) Then .CompanyID = SQLDataReader("CompanyID")
        '                If Not IsDBNull(SQLDataReader("CompanyName")) Then .CompanyName = SQLDataReader("CompanyName")
        '                If Not IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
        '                If Not IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
        '                If Not IsDBNull(SQLDataReader("FirstCategoryID")) Then .ProductCategory1 = SQLDataReader("FirstCategoryID")
        '                If Not IsDBNull(SQLDataReader("FirstCategoryName")) Then .ProductCategory1Name = SQLDataReader("FirstCategoryName")
        '                If Not IsDBNull(SQLDataReader("SecondCategoryID")) Then .ProductCategory2 = SQLDataReader("SecondCategoryID")
        '                If Not IsDBNull(SQLDataReader("SecondCategoryName")) Then .ProductCategory2Name = SQLDataReader("SecondCategoryName")

        '            End With
        '        Else
        '            bSuccess = False
        '            _strErrorMessage = "Product Series Not Found"                   'Set the Classes Error Message
        '            _intErrorNumber = 0 'Set the Classes Error Number

        '        End If
        '        'Cleanup
        '        SQLDataReader.Close()
        '        SQLConn.Close()



        '    Catch SQLErr As SqlException
        '        bSuccess = False
        '        _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
        '        _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

        '    Catch Err As Exception
        '        bSuccess = False
        '        _strErrorMessage = Err.ToString                      'Set the Classes Error Message
        '        _intErrorNumber = 0                                  'Set the Classes Error Number

        '    Finally
        '        'Confirm that The SQLDB Connection is closed
        '        If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
        '        SQLDataReader = Nothing
        '        SQLCommand = Nothing
        '        SQLConn = Nothing

        '    End Try

        '    Return bSuccess

        'End Function





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
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
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
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImagePath", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@Description").Value = sImage.Description
                SQLCommand.Parameters("@ImagePath").Value = sImage.ImagePath
                SQLCommand.Parameters("@ResourceTypeID").Value = sImage.ResourceTypeID
                SQLCommand.Parameters("@ResourceID").Value = sImage.ResourceID
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

                SQLCommand.CommandText = "LLFBPS..[GetImages]"

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

                SQLCommand.CommandText = "LLFBPS..[UpdateImageDescription]"

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
                SQLCommand.CommandText = "LLFBPS..[AddImage]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sImage" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImagePath", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))



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
        Public Function AddImageToResource(ByVal ResourceTypeID As Integer, ByVal ResourceID As Integer, ByVal ImageId As Integer) As Boolean

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
                SQLCommand.CommandText = "LLFBPS..[AddImageToResource]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sImage" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceTypeID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ImageID", System.Data.SqlDbType.Int, 4))


                SQLCommand.Parameters("@ResourceTypeID").Value = ResourceTypeID
                SQLCommand.Parameters("@ResourceID").Value = ResourceID
                SQLCommand.Parameters("@ImageID").Value = ImageId
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
                SQLCommand.CommandText = "LLFBPS..[SetNewPrimaryResourceID]"     'Stored Procedure Name

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
                SQLCommand.CommandText = "LLFBPS..[DeleteImageResourceJUNC]"     'Stored Procedure Name

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

#End Region
    End Class

    Public Class Order

        '"Order" Class Profile ******************************************************
        '
        '       Class:  Order
        '       Parent File: WorkElements.vb
        '       Parent Project: BPS
        '       Parent Solution: BPS
        '       Created By: Vincent Clover
        '       Created on: 02/20/06
        '       Description:  
        '
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


#Region "Private Variables"
        Private _lngOrderID As Long          'Used to Hold the OrderID currently being Used/Created
        Private _lngOrderLineItemID As Long          'Used to Hold the OrderID currently being Used/Created
        Private _intErrorNumber As Long        'Used to hold the ERROR code when Error occurs
        Private _strErrorMessage As String     'Used to hold the ERROR Message when Error occurs




#End Region

#Region "Class Properties"


        'Gets & Sets Current OrderID
        Public Property OrderID() As Long
            Get
                Return _lngOrderID  'Gets Local _ID Variable
            End Get
            Set(ByVal Value As Long)
                _lngOrderID = Value 'Sets Local _ID Variable
            End Set
        End Property


        ' Returns Error Message (When Error Occurs)
        Public ReadOnly Property ErrorMessage() As String
            Get
                Return _strErrorMessage
            End Get
        End Property


        ' Returns Error Number (When Error Occurs)
        Public ReadOnly Property ErrorNumber() As Long
            Get
                Return _intErrorNumber
            End Get
        End Property


#End Region

#Region "Constructors"


        'Empty Consructor
        Public Sub New()

        End Sub

        'Use this Constructor when wanting to set the OrderID on Creation
        Public Sub New(ByVal OrderID As Integer)

            _lngOrderID = OrderID 'Sets Local ID Variable

        End Sub



        ' Use this Constructor when wanting to create a NEW Order(Requires a "sOrder") 
        Public Sub New(ByRef sOrder As sCustomerOrder, ByRef CreatedOrder As Boolean)

            '  CreatedOrder = AddOrder(sCustomerOrder)

        End Sub




#End Region


#Region "Functions & Procedures"

        Public Function GetOrder(ByRef sCustomerOrder As sCustomerOrder) As Boolean
            _lngOrderID = sCustomerOrder.OrderID
            Return _GetOrder(sCustomerOrder)


        End Function

        Public Function GetOrder(ByRef sCustomerOrder As sCustomerOrder, ByVal OrderID As Long) As Boolean
            _lngOrderID = OrderID
            Return _GetOrder(sCustomerOrder)
        End Function

        

        Public Function GetOrder(ByRef sCustomerOrder As sCustomerOrder, ByVal OrderNumber As String, ByVal bIsOrderNumber As Boolean) As Boolean
            'Lookup OrderID by Order Number
            _lngOrderID = _GetOrderIDBYOrderNumber(OrderNumber)

            If _lngOrderID > 0 Then
                Return _GetOrder(sCustomerOrder)
            Else
                Return False
            End If

        End Function

        Public Function GetOrderLineItem(ByRef sCustomerOrderLineItem As sCustomerOrderLineItem) As Boolean
            _lngOrderLineItemID = sCustomerOrderLineItem.OrderLineItemID
            Return _GetOrderLineItem(sCustomerOrderLineItem)


        End Function

        Public Function GetOrderLineItem(ByRef sCustomerOrderLineItem As sCustomerOrderLineItem, ByVal OrderLineItemID As Long) As Boolean
            _lngOrderLineItemID = OrderLineItemID
            Return _GetOrderLineItem(sCustomerOrderLineItem)
        End Function

       

        Private Function _GetOrder(ByRef sCustomerOrder As sCustomerOrder) As Boolean

            'Function: _GetOrder ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: _GetOrder
            '       Parent Class:  Order
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sCustomerOrder Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  02/20/06
            '       Description: Gets a Order
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try

                SQLConn.Open() 'Open Database

                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                          'Set the Connection

                SQLCommand.CommandText = "LLFBPS..[GetCustomerOrder]"

                'Set the ProdcutID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngOrderID))


                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    With sCustomerOrder

                        'Fill the Order Information
                        If Not IsDBNull(SQLDataReader("Active")) Then .Active = SQLDataReader("Active")
                        If Not IsDBNull(SQLDataReader("CompanyID")) Then .CompanyID = SQLDataReader("CompanyID")
                        If Not IsDBNull(SQLDataReader("CompanyName")) Then .CompanyName = SQLDataReader("CompanyName")
                        If Not IsDBNull(SQLDataReader("BrandID")) Then .BrandID = SQLDataReader("BrandID")
                        If Not IsDBNull(SQLDataReader("BrandName")) Then .BrandName = SQLDataReader("BrandName")
                        If Not IsDBNull(SQLDataReader("CreatedBy")) Then .CreatedBy = SQLDataReader("CreatedBy")
                        If Not IsDBNull(SQLDataReader("CustomerID")) Then .CustomerID = SQLDataReader("CustomerID")
                        If Not IsDBNull(SQLDataReader("CaseID")) Then .CaseID = SQLDataReader("CaseID")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("DateClosed")) Then .DateClosed = SQLDataReader("DateClosed")
                        If Not IsDBNull(SQLDataReader("DateUpdated")) Then .DateUpdated = SQLDataReader("DateUpdated")
                        If Not IsDBNull(SQLDataReader("Description")) Then .Description = SQLDataReader("Description")
                        If Not IsDBNull(SQLDataReader("ItemQuantity")) Then .ItemQuantity = SQLDataReader("ItemQuantity")
                        If Not IsDBNull(SQLDataReader("OrderID")) Then .OrderID = SQLDataReader("OrderID")
                        If Not IsDBNull(SQLDataReader("OrderNumber")) Then .OrderNumber = SQLDataReader("OrderNumber")
                        If Not IsDBNull(SQLDataReader("OrderSource")) Then .OrderSource = SQLDataReader("OrderSource")
                        If Not IsDBNull(SQLDataReader("SubTotal")) Then .SubTotal = SQLDataReader("SubTotal")
                        If Not IsDBNull(SQLDataReader("ShippingHandling")) Then .ShippingHandling = SQLDataReader("ShippingHandling")
                        If Not IsDBNull(SQLDataReader("AdditionalSH")) Then .AdditionalSH = SQLDataReader("AdditionalSH")
                        If Not IsDBNull(SQLDataReader("Tax")) Then .Tax = SQLDataReader("Tax")
                        If Not IsDBNull(SQLDataReader("OrderTotal")) Then .OrderTotal = SQLDataReader("OrderTotal")
                        If Not IsDBNull(SQLDataReader("TrackingNumber")) Then .TrackingNumber = SQLDataReader("TrackingNumber")
                        If Not IsDBNull(SQLDataReader("ManualOrder")) Then .ManualOrder = SQLDataReader("ManualOrder")
                        If Not IsDBNull(SQLDataReader("TrackingNumber")) Then .TrackingNumber = SQLDataReader("TrackingNumber")
                        If Not IsDBNull(SQLDataReader("ShoppingCart")) Then .ShoppingCart = SQLDataReader("ShoppingCart")
                        If Not IsDBNull(SQLDataReader("CreditCardInformation")) Then .CreditCardInformation = SQLDataReader("CreditCardInformation")
                        If Not IsDBNull(SQLDataReader("BillingAddress")) Then .BillingAddress = SQLDataReader("BillingAddress")
                        If Not IsDBNull(SQLDataReader("ShippingAddress")) Then .ShippingAddress = SQLDataReader("ShippingAddress")
                        If Not IsDBNull(SQLDataReader("PONumber")) Then .PONumber = SQLDataReader("PONumber")
                        If Not IsDBNull(SQLDataReader("StoreNumber")) Then .StoreNumber = SQLDataReader("StoreNumber")

                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Order Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If

                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()

                'If the Order was found Then get the Customer Information
                If bSuccess = True Then
                    Dim objCustomer As New Customer(sCustomerOrder.CustomerID)
                    Dim sCustomer As sCustomer = Nothing

                    Try
                        If objCustomer.GetCustomer(sCustomer) = True Then
                            sCustomerOrder.Customer = sCustomer
                        End If

                    Catch Err As Exception
                        bSuccess = False
                        _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                        _intErrorNumber = 0                                  'Set the Classes Error Number

                    Finally
                        objCustomer = Nothing
                    End Try
                End If



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

        Private Function _GetOrderIDBYOrderNumber(ByVal OrderNumber As String) As Integer

            'Function: _GetOrder ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: _GetOrder
            '       Parent Class:  Order
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sCustomerOrder Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  02/20/06
            '       Description: Gets a Order
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim OrderID As Integer

            Try

                SQLConn.Open() 'Open Database

                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                          'Set the Connection

                SQLCommand.CommandText = "LLFBPS..[GetCustomerOrderByOrderNumber]"

                'Set the ProdcutID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderNumber", System.Data.SqlDbType.VarChar, 50, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, OrderNumber))


                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Product Structure

                If SQLDataReader.Read Then

                    OrderID = SQLDataReader("OrderID")
                Else
                    OrderID = 0
                    _strErrorMessage = "Order Not Found"                   'Set the Classes Error Message
                    _intErrorNumber = 0 'Set the Classes Error Number

                End If

                'Cleanup
                SQLDataReader.Close()
                SQLConn.Close()


            Catch SQLErr As SqlException
                OrderID = 0
                _strErrorMessage = SQLErr.ToString                   'Set the Classes Error Message
                _intErrorNumber = SQLErr.Number                      'Set the Classes Error Number

            Catch Err As Exception
                OrderID = 0
                _strErrorMessage = Err.ToString                      'Set the Classes Error Message
                _intErrorNumber = 0                                  'Set the Classes Error Number

            Finally
                'Confirm that The SQLDB Connection is closed
                If SQLConn.State = ConnectionState.Open Then SQLConn.Close()
                SQLDataReader = Nothing
                SQLCommand = Nothing
                SQLConn = Nothing

            End Try

            'Set returning OrderID
            Return OrderID


        End Function

        Private Function _GetOrderLineItem(ByRef sCustomerOrderLineItem As sCustomerOrderLineItem) As Boolean

            'Function: _GetOrderLineItem ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: _GetOrderLineItem
            '       Parent Class:  Order
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sCustomerOrderLineItem Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  11/15/06
            '       Description: Gets a OrderLineItem
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean = True 'Used to confirm that the Function Succesfully Processed

            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object

            Dim SQLCommand As New SqlCommand   'SQLCommand Object
            Dim SQLDataReader As SqlDataReader 'SQLDataReader
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            Try 'Two "Try" statements 

                SQLConn.Open() 'Open Database

                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure     'We're using Stored Procedures
                SQLCommand.Connection = SQLConn                          'Set the Connection

                SQLCommand.CommandText = "LLFBPS..[GetCustomerOrderLineItem]"

                'Set the ProdcutID Parameter
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderLineItemID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, _lngOrderLineItemID))


                'Set the DataReader
                SQLDataReader = SQLCommand.ExecuteReader(CommandBehavior.SingleResult)


                'Populate the Structure

                If SQLDataReader.Read Then

                    With sCustomerOrderLineItem

                        'Fill the OrderLineItem Information

                        If Not IsDBNull(SQLDataReader("BpsID")) Then .BpsID = SQLDataReader("BpsID")
                        If Not IsDBNull(SQLDataReader("BpsItemType")) Then .BpsItemType = SQLDataReader("BpsItemType")
                        If Not IsDBNull(SQLDataReader("CaseID")) Then .CaseID = SQLDataReader("CaseID")
                        If Not IsDBNull(SQLDataReader("DateCreated")) Then .DateCreated = SQLDataReader("DateCreated")
                        If Not IsDBNull(SQLDataReader("ItemNumber")) Then .ItemNumber = SQLDataReader("ItemNumber")
                        If Not IsDBNull(SQLDataReader("ItemDescription")) Then .ItemDescription = SQLDataReader("ItemDescription")
                        If Not IsDBNull(SQLDataReader("LastUpdated")) Then .LastUpdated = SQLDataReader("LastUpdated")
                        If Not IsDBNull(SQLDataReader("OrderID")) Then .OrderID = SQLDataReader("OrderID")
                        If Not IsDBNull(SQLDataReader("OrderLineItemID")) Then .OrderLineItemID = SQLDataReader("OrderLineItemID")
                        If Not IsDBNull(SQLDataReader("Price")) Then .Price = SQLDataReader("Price")
                        If Not IsDBNull(SQLDataReader("Tax")) Then .Tax = SQLDataReader("Tax")
                        If Not IsDBNull(SQLDataReader("ProblemModel")) Then .ProblemModel = SQLDataReader("ProblemModel")
                        If Not IsDBNull(SQLDataReader("Quantity")) Then .Quantity = SQLDataReader("Quantity")
                        If Not IsDBNull(SQLDataReader("ReasonCode")) Then .ReasonCode = SQLDataReader("ReasonCode")
                        If Not IsDBNull(SQLDataReader("ReasonCodeDescription")) Then .ReasonCodeDescription = SQLDataReader("ReasonCodeDescription")
                        If Not IsDBNull(SQLDataReader("Status")) Then .Status = SQLDataReader("Status")
                        If Not IsDBNull(SQLDataReader("StatusDescription")) Then .StatusDescription = SQLDataReader("StatusDescription")
                        If Not IsDBNull(SQLDataReader("TrackingNumber")) Then .TrackingNumber = SQLDataReader("TrackingNumber")
                        If Not IsDBNull(SQLDataReader("TrackingSource")) Then .TrackingSource = SQLDataReader("TrackingSource")
                        If Not IsDBNull(SQLDataReader("UnderWarranty")) Then .UnderWarranty = SQLDataReader("UnderWarranty")
                        If Not IsDBNull(SQLDataReader("ShippingAddress")) Then .ShippingAddress = SQLDataReader("ShippingAddress")
                        If Not IsDBNull(SQLDataReader("DateShipped")) Then .DateShipped = SQLDataReader("DateShipped")


                    End With
                Else
                    bSuccess = False
                    _strErrorMessage = "Order Line Not Found"                   'Set the Classes Error Message
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






        Public Function AddCustomerOrder(ByRef sCustomerOrder As sCustomerOrder) As Boolean


            'Function: AddCustomerOrder ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddCustomerOrder
            '       Parent Class:  Order
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sCustomerOrder Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Amy Cook
            '       Created on:  02/28/06
            '       Description: Adds an entry to the CustomerOrders table
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add a Customer Order One Step is Required: ********************
            '   - 1. Add an entry to the CustomerOrders table
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[AddCustomerOrder]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomerOrder" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CompanyID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CustomerID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CaseID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderNumber", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@SubTotal", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShippingHandling", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@AdditionalSH", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Tax", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderTotal", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ItemQuantity", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreatedBy", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderSource", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ManualOrder", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Description", System.Data.SqlDbType.VarChar, 1500))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TrackingNumber", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Active", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShoppingCart", System.Data.SqlDbType.Text, 6000))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CreditCardInformation", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BillingAddress", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShippingAddress", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@PONumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@StoreNumber", System.Data.SqlDbType.VarChar, 50))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))




                'Set the Values

                SQLCommand.Parameters("@CompanyID").Value = sCustomerOrder.CompanyID
                SQLCommand.Parameters("@BrandID").Value = sCustomerOrder.BrandID
                SQLCommand.Parameters("@CustomerID").Value = sCustomerOrder.CustomerID
                SQLCommand.Parameters("@CaseID").Value = sCustomerOrder.CaseID
                SQLCommand.Parameters("@OrderNumber").Value = sCustomerOrder.OrderNumber
                SQLCommand.Parameters("@SubTotal").Value = sCustomerOrder.SubTotal
                SQLCommand.Parameters("@ShippingHandling").Value = sCustomerOrder.ShippingHandling
                SQLCommand.Parameters("@AdditionalSH").Value = sCustomerOrder.AdditionalSH
                SQLCommand.Parameters("@Tax").Value = sCustomerOrder.Tax
                SQLCommand.Parameters("@OrderTotal").Value = sCustomerOrder.OrderTotal
                SQLCommand.Parameters("@ItemQuantity").Value = sCustomerOrder.ItemQuantity
                SQLCommand.Parameters("@CreatedBy").Value = sCustomerOrder.CreatedBy
                SQLCommand.Parameters("@OrderSource").Value = sCustomerOrder.OrderSource
                SQLCommand.Parameters("@ManualOrder").Value = sCustomerOrder.ManualOrder
                SQLCommand.Parameters("@Description").Value = sCustomerOrder.Description
                SQLCommand.Parameters("@TrackingNumber").Value = sCustomerOrder.TrackingNumber
                SQLCommand.Parameters("@Active").Value = sCustomerOrder.Active
                SQLCommand.Parameters("@ShoppingCart").Value = sCustomerOrder.ShoppingCart
                SQLCommand.Parameters("@CreditCardInformation").Value = sCustomerOrder.CreditCardInformation
                SQLCommand.Parameters("@BillingAddress").Value = sCustomerOrder.BillingAddress
                SQLCommand.Parameters("@ShippingAddress").Value = sCustomerOrder.ShippingAddress
                SQLCommand.Parameters("@PONumber").Value = sCustomerOrder.PONumber
                SQLCommand.Parameters("@StoreNumber").Value = sCustomerOrder.StoreNumber



                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "OrderID" back to the sCustomerOrder Structure
                sCustomerOrder.OrderID = SQLCommand.Parameters("@OrderID").Value

                If Not SQLCommand.Parameters("@OrderID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Customer Order (Not System Error)"     'Set the Classes Error Message
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



        Public Function AddCustomerOrderLineItem(ByRef sCustomerOrderLineItem As sCustomerOrderLineItem) As Boolean


            'Function: AddCustomerOrderLineItem ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddCustomerOrderLineItem
            '       Parent Class:  Order
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sCustomerOrderLineItem Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  11/01/06
            '       Description: Adds an entry to the CustomerOrderLineItem table
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To Add a Customer Order Line item One Step is Required: ********************
            '   - 1. Add an entry to the CustomerOrders table
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[AddCustomerOrderLineItem]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomerOrderLineItem" Structure


                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BpsItemType", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CaseID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ItemNumber", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ItemDescription", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BpsID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ReasonCode", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ReasonCodeDescription", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProblemModel", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Tax", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Price", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Quantity", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UnderWarranty", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Status", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@StatusDescription", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TrackingNumber", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TrackingSource", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShippingAddress", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderLineItemID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))


                SQLCommand.Parameters("@OrderID").Value = sCustomerOrderLineItem.OrderID
                SQLCommand.Parameters("@BpsItemType").Value = sCustomerOrderLineItem.BpsItemType
                SQLCommand.Parameters("@CaseID").Value = sCustomerOrderLineItem.CaseID
                SQLCommand.Parameters("@ItemNumber").Value = sCustomerOrderLineItem.ItemNumber
                SQLCommand.Parameters("@ItemDescription").Value = sCustomerOrderLineItem.ItemDescription
                SQLCommand.Parameters("@BpsID").Value = sCustomerOrderLineItem.BpsID
                SQLCommand.Parameters("@ReasonCode").Value = sCustomerOrderLineItem.ReasonCode
                SQLCommand.Parameters("@ReasonCodeDescription").Value = sCustomerOrderLineItem.ReasonCodeDescription
                SQLCommand.Parameters("@ProblemModel").Value = sCustomerOrderLineItem.ProblemModel
                SQLCommand.Parameters("@Price").Value = sCustomerOrderLineItem.Price
                SQLCommand.Parameters("@Tax").Value = sCustomerOrderLineItem.Tax
                SQLCommand.Parameters("@Quantity").Value = sCustomerOrderLineItem.Quantity
                SQLCommand.Parameters("@UnderWarranty").Value = sCustomerOrderLineItem.UnderWarranty
                SQLCommand.Parameters("@Status").Value = sCustomerOrderLineItem.Status
                SQLCommand.Parameters("@StatusDescription").Value = sCustomerOrderLineItem.StatusDescription
                SQLCommand.Parameters("@TrackingNumber").Value = sCustomerOrderLineItem.TrackingNumber
                SQLCommand.Parameters("@TrackingSource").Value = sCustomerOrderLineItem.TrackingSource
                SQLCommand.Parameters("@ShippingAddress").Value = sCustomerOrderLineItem.ShippingAddress


                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure

                'Pass the new "OrderID" back to the sCustomerOrder Structure
                sCustomerOrderLineItem.OrderLineItemID = SQLCommand.Parameters("@OrderLineItemID").Value

                If Not SQLCommand.Parameters("@OrderLineItemID").Value > 0 Then
                    _intErrorNumber = "Internal Error adding Customer Order Line Item (Not System Error)"     'Set the Classes Error Message
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












        Public Function UpdateCustomerOrderLineItem(ByRef sCustomerOrderLineItem As sCustomerOrderLineItem) As Boolean


            'Function: AddCustomerOrderLineItem ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            '
            '       Name: AddCustomerOrderLineItem
            '       Parent Class:  Order
            '       Parent File: WorkElements.vb
            '       Parent Project: BPS
            '       Parent Solution: BPS
            '
            '       Return Value: BOOLEAN
            '       Required Parameters(1): sCustomerOrderLineItem Structure(ByRef) 
            '       Optional Parameters(0):
            '
            '       Created By:  Vincent Clover
            '       Created on:  11/01/06
            '       Description: Adds an entry to the CustomerOrderLineItem table
            '
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim bSuccess As Boolean 'Return variable 

            'Set Local Variables '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString)

            Dim SQLCommand As New SqlCommand    'SQLCommand Object
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


            '*********************************************************************************************
            ' ****************** To update a Customer Order Line item One Step is Required: ********************
            '   - 1. Update the entry to the CustomerOrders table
            '   
            '*********************************************************************************************

            Try    'Only one "Try" statement 

                SQLConn.Open()    'Open Database


                'Set the Basic Command Information 
                SQLCommand.CommandType = CommandType.StoredProcedure    'We're using Stored Procedures
                SQLCommand.Connection = SQLConn       'Set the Connection


                'Set the Specific Command Information 
                SQLCommand.CommandText = "LLFBPS..[UpdateCustomerOrderLineItem]"     'Stored Procedure Name

                'Stored Procedure Paramaters - Set their Values using the "sCustomerOrderLineItem" Structure

                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderLineItemID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@OrderID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BpsItemType", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@CaseID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ItemNumber", System.Data.SqlDbType.VarChar, 150))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ItemDescription", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BpsID", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ReasonCode", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ReasonCodeDescription", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProblemModel", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Price", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Tax", System.Data.SqlDbType.Money, 10))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Quantity", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UnderWarranty", System.Data.SqlDbType.Bit, 1))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Status", System.Data.SqlDbType.Int, 4))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@StatusDescription", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TrackingNumber", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@TrackingSource", System.Data.SqlDbType.VarChar, 100))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShippingAddress", System.Data.SqlDbType.VarChar, 250))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@DateShipped", System.Data.SqlDbType.DateTime, 8))
                SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UpdatedDate", System.Data.SqlDbType.DateTime, 8, System.Data.ParameterDirection.Output, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))



                SQLCommand.Parameters("@OrderLineItemID").Value = sCustomerOrderLineItem.OrderLineItemID
                SQLCommand.Parameters("@OrderID").Value = sCustomerOrderLineItem.OrderID
                SQLCommand.Parameters("@BpsItemType").Value = sCustomerOrderLineItem.BpsItemType
                SQLCommand.Parameters("@CaseID").Value = sCustomerOrderLineItem.CaseID
                SQLCommand.Parameters("@ItemNumber").Value = sCustomerOrderLineItem.ItemNumber
                SQLCommand.Parameters("@ItemDescription").Value = sCustomerOrderLineItem.ItemDescription
                SQLCommand.Parameters("@BpsID").Value = sCustomerOrderLineItem.BpsID
                SQLCommand.Parameters("@ReasonCode").Value = sCustomerOrderLineItem.ReasonCode
                SQLCommand.Parameters("@ReasonCodeDescription").Value = sCustomerOrderLineItem.ReasonCodeDescription
                SQLCommand.Parameters("@ProblemModel").Value = sCustomerOrderLineItem.ProblemModel
                SQLCommand.Parameters("@Price").Value = sCustomerOrderLineItem.Price
                SQLCommand.Parameters("@Tax").Value = sCustomerOrderLineItem.Tax
                SQLCommand.Parameters("@Quantity").Value = sCustomerOrderLineItem.Quantity
                SQLCommand.Parameters("@UnderWarranty").Value = sCustomerOrderLineItem.UnderWarranty
                SQLCommand.Parameters("@Status").Value = sCustomerOrderLineItem.Status
                SQLCommand.Parameters("@StatusDescription").Value = sCustomerOrderLineItem.StatusDescription
                SQLCommand.Parameters("@TrackingNumber").Value = sCustomerOrderLineItem.TrackingNumber
                SQLCommand.Parameters("@TrackingSource").Value = sCustomerOrderLineItem.TrackingSource
                SQLCommand.Parameters("@ShippingAddress").Value = sCustomerOrderLineItem.ShippingAddress
                If sCustomerOrderLineItem.DateShipped > CDate("01/01/2006") Then
                    SQLCommand.Parameters("@DateShipped").Value = sCustomerOrderLineItem.DateShipped
                End If



                SQLCommand.ExecuteNonQuery()    'Execute the Stored Procedure


                If Not IsDate(SQLCommand.Parameters("@UpdatedDate").Value) Then
                    _intErrorNumber = "Internal Error Updating Customer Order Line Item (Not System Error)"     'Set the Classes Error Message
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













#End Region


    End Class



End Namespace