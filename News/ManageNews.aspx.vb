
Imports System.IO

Partial Class Media_ManageNews
    Inherits System.Web.UI.Page


    Private Sub DeleteFile(ByVal FilePath As String)
        Try
            File.Delete(FilePath)

        Catch ex As Exception

        End Try
    End Sub

    Protected Sub gvNews_RowDeleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeletedEventArgs) Handles gvNews.RowDeleted
        Try
            'Delete the File also 
            If ViewState("DeleteFile") <> "" Then
                DeleteFile(ViewState("DeleteFile"))
            End If

            'reset count total minus current record deleted...data is not rebinded...
            lblCount.Text = String.Format("({0})", gvNews.Rows.Count - 1)


        Catch ex As Exception

        End Try
    End Sub


  

    Protected Sub gvNews_RowDeleting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewDeleteEventArgs) Handles gvNews.RowDeleting
        Try
            ViewState("DeleteFile") = CType(CType(sender, GridView).Rows(e.RowIndex).FindControl("lblFilePath"), Label).Text


        Catch ex As Exception

        End Try
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not Page.IsPostBack Then
            lblCount.Text = String.Format("({0})", gvNews.Rows.Count)
            Me.Master.PageHeader = "Manage News"
            Me.Master.PageTitle = "Manage News"
            Me.Master.SetDisplayMessage("Use the links below to Add, Edit or Remove Quantum News Items", Management.MessageType.GeneralMessage)


        End If
    End Sub


End Class
