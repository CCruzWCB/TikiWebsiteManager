Imports AspCrypt.CryptClass
Imports System.Data.SqlClient
Imports System.Web.Security

Namespace Management
    Public Class Security
        Dim _UserID As Integer = 0
        Dim _ErrorNumber As String
        Dim _Name As String
        Dim _ErrorDescription As String

        Public ReadOnly Property UserID() As Integer
            Get
                Return _UserID
            End Get
        End Property

        Public ReadOnly Property Name() As String
            Get
                Return _Name
            End Get
        End Property

        Public ReadOnly Property ErrorDescription()
            Get
                Return _ErrorDescription
            End Get
        End Property
        Public ReadOnly Property ErrorNumber()
            Get
                Return _ErrorNumber
            End Get
        End Property


        Private Function Encrypt(ByVal Value As String, ByVal key As String) As String
            Dim oCrypt As New AspCrypt.Crypt
            Return oCrypt.Crypt(UCase(Value), UCase(key))
            oCrypt = Nothing
        End Function

        Public Function Authenticated(ByVal UserName As String, ByVal Password As String) As Boolean

            Dim SqlConn As New SqlConnection
            Dim CmdAuthenticate As New SqlCommand
            '
            'SqlConn
            '   
            SqlConn.ConnectionString = ConfigurationManager.AppSettings("BradleyIDConnectionString").ToString

            '
            'CmdAuthenticate
            '
            CmdAuthenticate.CommandText = "dbo.[sp_AuthenticateBradleyIDByUserName]"
            CmdAuthenticate.CommandType = System.Data.CommandType.StoredProcedure
            CmdAuthenticate.Connection = SqlConn
            CmdAuthenticate.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
            CmdAuthenticate.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserName", System.Data.SqlDbType.VarChar, 100))

            Dim encPassword As String

            Try
                CmdAuthenticate.Connection.Open()
                CmdAuthenticate.Parameters("@UserName").Value = UserName
                Dim oReader As SqlDataReader = CmdAuthenticate.ExecuteReader

                If oReader.Read() Then
                    _UserID = oReader("Member_ID")
                    encPassword = oReader("Password")
                    _Name = oReader("Description")
                    Dim tempEncPassword As String = Encrypt(UserName, Password)

                    If encPassword = tempEncPassword Then
                        _ErrorDescription = "User Authenticated!"
                    Else
                        _UserID = 0
                        _ErrorDescription = "User Password is Incorrect!"
                    End If
                Else
                    _ErrorDescription = "Invalid User Name, Please Try Again."
                End If

            Catch e As SqlException
                _ErrorDescription = e.Message.ToString
                _UserID = 0

            Catch ex As Exception
                _UserID = 0
                _ErrorDescription = ex.Message.ToString
            Finally
                If SqlConn.State = Data.ConnectionState.Open Then SqlConn.Close()
            End Try

            Return UserID

        End Function

        Public Function GetAccessType(ByVal UserID As Integer, ByVal Resource As ResourceSectionType) As AccessType

            Dim SqlCnn As New SqlConnection
            Dim CmdGetAccess As New System.Data.SqlClient.SqlCommand
            Dim oReader As SqlDataReader

            SqlCnn.ConnectionString = ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString

            CmdGetAccess.CommandText = "dbo.[GetAccessTypeByResourceID]"
            CmdGetAccess.CommandType = System.Data.CommandType.StoredProcedure
            CmdGetAccess.Connection = SqlCnn
            CmdGetAccess.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
            CmdGetAccess.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserID", System.Data.SqlDbType.Int, 4))
            CmdGetAccess.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
            CmdGetAccess.Parameters("@UserID").Value = UserID
            CmdGetAccess.Parameters("@ResourceID").Value = CType(Resource, Integer)

            Try
                SqlCnn.Open()
                oReader = CmdGetAccess.ExecuteReader

                If oReader.Read Then
                    Return CType(oReader("AccessType_ID"), AccessType)
                Else
                    Return AccessType.NoAccess
                End If

            Catch e As SqlException
                _ErrorDescription = e.Message.ToString
                Return AccessType.NoAccess
            Catch ex As Exception
                _ErrorDescription = ex.Message.ToString
                Return AccessType.NoAccess
            Finally
                If SqlCnn.State = Data.ConnectionState.Open Then SqlCnn.Close()
            End Try


        End Function

        Public Function GetAccessTypeByResourceID(ByVal UserID As Integer, ByVal ResourceID As Integer) As AccessType
            Dim SqlCnn As New SqlConnection
            Dim CmdGetAccess As New System.Data.SqlClient.SqlCommand
            Dim oReader As SqlDataReader

            SqlCnn.ConnectionString = ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString

            CmdGetAccess.CommandText = "dbo.[GetAccessTypeByResourceID]"
            CmdGetAccess.CommandType = System.Data.CommandType.StoredProcedure
            CmdGetAccess.Connection = SqlCnn
            CmdGetAccess.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
            CmdGetAccess.Parameters.Add(New System.Data.SqlClient.SqlParameter("@UserID", System.Data.SqlDbType.Int, 4))
            CmdGetAccess.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ResourceID", System.Data.SqlDbType.Int, 4))
            CmdGetAccess.Parameters("@UserID").Value = UserID
            CmdGetAccess.Parameters("@ResourceID").Value = ResourceID

            Try
                SqlCnn.Open()
                oReader = CmdGetAccess.ExecuteReader

                If oReader.Read Then
                    Return CType(oReader("AccessType_ID"), AccessType)
                Else
                    Return AccessType.NoAccess
                End If

            Catch e As SqlException
                _ErrorDescription = e.Message.ToString
                Return AccessType.NoAccess
            Catch ex As Exception
                _ErrorDescription = ex.Message.ToString
                Return AccessType.NoAccess
            Finally
                If SqlCnn.State = Data.ConnectionState.Open Then SqlCnn.Close()
            End Try

        End Function

        Public Enum AccessType
            NoAccess = 1
            _ReadOnly = 2
            _FullAccess = 3
            ' _ReservedAccessType1 = 4
            ' _ReservedAccessType2 = 5
            ' _ReservedAccessType3 = 6
        End Enum


        Public Enum ResourceSectionType
            OrdersSection = 1
            Customers = 2
            ItemManagement = 3
            PartsManagement = 4
            DocumentManagement = 5
            SourceCodeReferrels = 6
            GiftRegistrySection = 7
            RecipeManagement = 8
            RetailerManagement = 9
            WebSiteDesignManagement = 10
            PromotionsManagement = 11
            DataCubesManagement = 12
            GroupItemsManagement = 13
            ManagerLandingPage = 14
            UserManagmentSection = 15
            MetaData = 16
            ManualsManagement = 17
            ProductsManagement = 18
            ProductsCategoryManagement = 19
            DocumentCategoryManagement = 20
            SpecificationManagement = 21
            WarrantyManagement = 22
            GrillTypes = 23

        End Enum

    End Class

End Namespace
