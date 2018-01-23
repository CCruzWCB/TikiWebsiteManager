
Partial Class Media_SearchRedirects
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Set the default focus
        txtName.Focus()

        If Not Page.IsPostBack Then

            Me.Master.PageHeader = "Manage Search Redirects"
            Me.Master.PageTitle = "Manage Search Redirects"
            Me.Master.SetDisplayMessage("Search Redirects are used to control what page a user will be redirected to when searching for select search terms.<br>  Use the form below to Add, Edit or Remove Search Redirects.", Management.MessageType.GeneralMessage)

            'Set the results count
            If Me.gvSearchRedirect.Rows.Count > 0 Then
                lblCount.Text = String.Format("({0})", gvSearchRedirect.Rows.Count)


            End If
        End If
    End Sub

    Protected Sub dsSearchRedirect_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsSearchRedirect.Inserted
        Try
            'Check for errors and display INSERT status
            If IsNothing(e.Exception) Then
                Me.Master.SetDisplayMessage("Search Redirect INSERT Successful", Management.MessageType.ConfirmationMessage)
            Else
                Me.Master.SetDisplayMessage("Error INSERTING Search Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)
            End If

        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error INSERTING Search Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)

        End Try

    End Sub

    Protected Sub dsSearchRedirect_Updated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsSearchRedirect.Updated
        Try
            'Check for errors and display UPDATE status
            If IsNothing(e.Exception) Then
                Me.Master.SetDisplayMessage("Search Redirect UPDATE Successful", Management.MessageType.ConfirmationMessage)
            Else
                Me.Master.SetDisplayMessage("Error UPDATING Search Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)
            End If

        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error UPDATING Search Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)

        End Try

    End Sub


    Protected Sub btnAddNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddNew.Click

        Try
            'Set Insert Paramaters
            dsSearchRedirect.InsertParameters("Name").DefaultValue = txtName.Text
            dsSearchRedirect.InsertParameters("RedirectURL").DefaultValue = txtURL.Text


            If IsNumeric(Page.User.Identity.Name) Then
                dsSearchRedirect.InsertParameters("CreatedBy").DefaultValue = Page.User.Identity.Name
            Else
                dsSearchRedirect.InsertParameters("CreatedBy").DefaultValue = 0
            End If



            dsSearchRedirect.InsertParameters("Active").DefaultValue = chkActive.Checked

            'Insert Record
            dsSearchRedirect.Insert()

            'Rebind the datagrid
            Me.gvSearchRedirect.DataBind()
        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error INSERTING Search Redirect:  " & ex.Message.ToString, Management.MessageType.ErrorMessage)
        End Try
      
    End Sub
End Class
