
Imports System.Data.SqlClient
Imports System.Web.UI.WebControls
Imports System.Data
Imports BPS_BL.BPS

Namespace Management
    Namespace TrickBag

      


        Public NotInheritable Class Convert
            Public Shared Function ArrayListToDataSet(ByVal strtableName As String, ByVal strColumnName As String, ByVal arrList As ArrayList) As Data.DataSet
                Dim ds As New Data.DataSet
                Dim dr As DataRow
                Dim i As Long

                ds.tables.Add(strtableName)
                ds.tables(strtableName).Columns.Add(strColumnName)

                For i = 0 To arrList.Count - 1
                    dr = ds.tables(strtableName).NewRow
                    dr.Item(strColumnName) = arrList.Item(i)
                    ds.tables(strtableName).Rows.Add(dr)
                Next
                Return ds
            End Function

            Public Shared Function DataReaderToDataSet(ByVal dr As SqlDataReader) As DataSet
                Dim ds As New DataSet
                Dim i As Integer
                Dim dRow As DataRow
                Dim c As DataColumn

                Do
                    'Create New Data table
                    Dim schematable = dr.GetSchematable()
                    Dim dt As New Datatable

                    'If (schematable Is Nothing) Then

                    For i = 0 To schematable.Rows.Count - 1
                        dRow = schematable.Rows(i)
                        c = New DataColumn(dRow("ColumnName"), dRow("DataType"))
                        dt.Columns.Add(c)
                    Next
                    ds.tables.Add(dt)

                    Do While (dr.Read())
                        dRow = dt.NewRow()
                        For i = 0 To dr.FieldCount - 1
                            dRow(i) = dr.GetValue(i)
                        Next
                        dt.Rows.Add(dRow)
                    Loop
                    'Else
                    'c = New DataColumn("RowsAffected")
                    'dt.Columns.Add(c)
                    'ds.tables.Add(dt)
                    'dRow = dt.NewRow()
                    'dRow(0) = dr.RecordsAffected
                    'dt.Rows.Add(dRow)
                    'End If
                Loop While (dr.NextResult())
                Return ds
            End Function

        End Class


        Public NotInheritable Class Add
            Public Shared Function ItemtoDropDownList(ByVal ItemText As String, ByVal ItemValue As String, Optional ByVal ItemSelected As Boolean = False)

                Dim TempListItem As New ListItem
                With TempListItem
                    .Text = ItemText
                    .Value = ItemValue
                    .Selected = ItemSelected 'True
                End With
                Return TempListItem
            End Function

         
        End Class
    End Namespace
End Namespace