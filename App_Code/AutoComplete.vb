Imports System.Web
Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.Data.SqlClient
Imports System.Web.Script.Services

<WebService(Namespace:="http://tempuri.org/")> _
<WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ScriptService()> _
<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Public Class AutoComplete
    Inherits System.Web.Services.WebService


    <WebMethod()> _
    Public Function GetModelCompletionList(ByVal prefixText As String, ByVal count As Integer) As String()

        Dim tempString() As String = Nothing


        tempString = GetItems(prefixText, count, ReturnTypeEnum.CommaDelimatedList)

        Return tempString

    End Function

    '<WebMethod()> _
    'Public Function GetDealerCompletionList(ByVal prefixText As String) As String()

    '    Dim tempString() As String = Nothing


    '    tempString = GetBusinessCustomers(prefixText, 2, ReturnTypeEnum.CommaDelimatedList)

    '    Return tempString

    'End Function

    'Public Function GetBusinessCustomers(ByVal PrefixText As String, ByVal Count As Integer, Optional ByVal myReturnTypeEnum As ReturnTypeEnum = ReturnTypeEnum.CommaDelimatedList) As String()
    '    Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object
    '    Dim SQLDR As SqlDataReader
    '    Dim SQLCommand As New SqlCommand
    '    Dim items As New ArrayList
    '    Dim ListArray() As String = Nothing

    '    Dim myresult As String = ""


    '    Try
    '        SQLCommand.CommandText = "Support.[GetDealersAutoCompleteListings]"
    '        SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
    '        SQLCommand.Connection = SQLConn
    '        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Prefix", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, PrefixText))
    '        SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShowList", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CType(myReturnTypeEnum, Boolean)))


    '        SQLConn.Open()



    '        SQLDR = SQLCommand.ExecuteReader



    '        Do While SQLDR.Read

    '            ListArray = SQLDR("Results").ToString.Split(",")

    '        Loop




    '    Catch SQLErr As SqlException

    '        'TODO create sys message for role not found 
    '    Catch ex As Exception

    '        'TODO create sys message for role not found 
    '    Finally


    '        If SQLConn.State = Data.ConnectionState.Open Then SQLConn.Close()


    '        SQLConn = Nothing
    '        SQLDR = Nothing
    '        SQLCommand = Nothing
    '    End Try

    '    '  items = myresult.ToString



    '    Return ListArray





    'End Function

    Public Function GetItems(ByVal PrefixText As String, ByVal Count As Integer, Optional ByVal myReturnTypeEnum As ReturnTypeEnum = ReturnTypeEnum.CommaDelimatedList) As String()
        Dim SQLConn As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDBConnection").ToString) 'SQLConnection Object
        Dim SQLDR As SqlDataReader
        Dim SQLCommand As New SqlCommand
        Dim items As New ArrayList
        Dim ListArray() As String = Nothing

        Dim myresult As String = ""


        Try
            SQLCommand.CommandText = "LLFBPS.dbo.[GetItemAutoCompleteListings]"
            SQLCommand.CommandType = System.Data.CommandType.StoredProcedure
            SQLCommand.Connection = SQLConn
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@Prefix", System.Data.SqlDbType.VarChar, 50, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, PrefixText))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ShowList", System.Data.SqlDbType.Bit, 1, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, CType(myReturnTypeEnum, Boolean)))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, ConfigurationManager.AppSettings("BrandID")))

            SQLConn.Open()



            SQLDR = SQLCommand.ExecuteReader



            Do While SQLDR.Read

                ListArray = SQLDR("Results").ToString.Split(",")

            Loop




        Catch SQLErr As SqlException

            'TODO create sys message for role not found 
        Catch ex As Exception

            'TODO create sys message for role not found 
        Finally


            If SQLConn.State = Data.ConnectionState.Open Then SQLConn.Close()


            SQLConn = Nothing
            SQLDR = Nothing
            SQLCommand = Nothing
        End Try

        '  items = myresult.ToString



        Return ListArray





    End Function

    Public Enum ReturnTypeEnum
        DataTable = 0
        CommaDelimatedList = 1

    End Enum
End Class