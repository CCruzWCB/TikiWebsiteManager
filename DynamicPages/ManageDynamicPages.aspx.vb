Imports Management
Partial Class FooterContent_ManageFooterContent
    Inherits System.Web.UI.Page

    

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.Master.PageHeader = "Manage Static Footer Content"
        Me.Master.PageTitle = "Manage Static Footer Content"
        Me.Master.SetDisplayMessage("Welcome to Manage Static Footer Content Management", MessageType.GeneralMessage)

    End Sub

    Protected Sub btnAdd_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        dsDocList.InsertParameters("doc_name").DefaultValue = Me.txtName.Text
        dsDocList.Insert()
    End Sub

    Protected Sub dsDocList_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsDocList.Inserted

        If IsNumeric(e.Command.Parameters("@doc_id").Value) Then
            ViewState("doc_id") = e.Command.Parameters("@doc_id").Value


            If ViewState("doc_id") > 0 Then
                Response.Redirect(String.Format("EditDocument.aspx?doc_id={0}", ViewState("doc_id")))

            End If
        Else
            Me.Master.SetDisplayMessage("Error Adding Part!  " & e.Exception.ToString, MessageType.ErrorMessage)
        End If
    End Sub
End Class
