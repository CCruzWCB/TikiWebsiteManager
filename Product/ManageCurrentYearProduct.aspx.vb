Imports System.Data
Imports System.Data.SqlClient
Imports Management

Partial Class Product_ManageCurrentYearProduct
    Inherits System.Web.UI.Page
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Master.PageHeader = "Manage Current Year Products"
        Me.Master.PageTitle = "Manage Current Year Products "
        Me.Master.SetDisplayMessage("Welcome to Current Year Products Management", MessageType.GeneralMessage)


    End Sub



    Protected Sub btnMoveRight_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMoveRight.Click

        Dim liNonMktgProd As ListItem
        Try

            For Each liNonMktgProd In Me.lbNonMktg.Items
                If liNonMktgProd.Selected Then
                    'Update Product Attributes flag
                    UpdateCurrentMKTGProduct(liNonMktgProd.Value, True)

                End If
            Next

            Me.lbNonMktg.DataBind()
            Me.lbMktg.DataBind()

        Catch ex As Exception
        Finally
            liNonMktgProd = Nothing
        End Try


    End Sub

    Public Function UpdateCurrentMKTGProduct(ByVal ProductID As Integer, ByVal IsActiveOnMktgSite As Boolean) As Boolean


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
            SQLCommand.CommandText = "LLFBPS..[MANAGERAddCurrentYearProduct]"     'Stored Procedure Name

            'Stored Procedure Paramaters - Set their Values using the "sKnowledge" Structure


            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@RETURN_VALUE", System.Data.SqlDbType.Int, 4, System.Data.ParameterDirection.ReturnValue, False, CType(0, Byte), CType(0, Byte), "", System.Data.DataRowVersion.Current, Nothing))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@ProductID", System.Data.SqlDbType.Int, 4))
            SQLCommand.Parameters.Add(New System.Data.SqlClient.SqlParameter("@IsCurrentYearProduct", System.Data.SqlDbType.Bit, 4))


            SQLCommand.Parameters("@ProductID").Value = ProductID
            SQLCommand.Parameters("@IsCurrentYearProduct").Value = IsActiveOnMktgSite

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

    Protected Sub btnMoveLeft_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnMoveLeft.Click
        Dim liMktgProd As ListItem
        Try

            For Each liMktgProd In Me.lbMktg.Items
                If liMktgProd.Selected Then
                    'Update Product Attributes flag
                    UpdateCurrentMKTGProduct(liMktgProd.Value, False)
                End If
            Next

            Me.lbNonMktg.DataBind()
            Me.lbMktg.DataBind()
        Catch ex As Exception
        Finally
            liMktgProd = Nothing
        End Try

    End Sub
End Class
