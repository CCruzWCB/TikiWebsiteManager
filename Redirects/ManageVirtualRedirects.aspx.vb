
Partial Class Features_ManageVirtualRedirects
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Set the default focus
        txtName.Focus()

        If Not Page.IsPostBack Then

            Me.Master.PageHeader = "Manage Virtual Folder Redirects"
            Me.Master.PageTitle = "Manage Virtual Folder Redirects"
            Me.Master.SetDisplayMessage("Virtual Folder Redirects are used to control virtual folders names used as shortcuts to internal pages. (Example: www.Char-Broil.com/MyVirtualFolder)  <br>  Use the form below to Add, Edit or Remove Virtual Redirects.", Management.MessageType.GeneralMessage)

            'Set the results count
            If Me.gvVirtualRedirect.Rows.Count > 0 Then
                lblCount.Text = String.Format("({0})", gvVirtualRedirect.Rows.Count)


            End If
        End If
    End Sub

    Protected Sub dsVirtualRedirect_Deleted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsVirtualRedirect.Deleted
        Try
            'Check for errors and display INSERT status
            If IsNothing(e.Exception) Then
                Me.Master.SetDisplayMessage("Virtual Redirect Deletion was Successful", Management.MessageType.ConfirmationMessage)
            Else
                Me.Master.SetDisplayMessage("Error Deleting Virtual Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)
            End If
        Catch ex As Exception

        End Try
    End Sub

    Protected Sub dsVirtualRedirect_Inserted(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsVirtualRedirect.Inserted
        Try
            'Check for errors and display INSERT status
            If IsNothing(e.Exception) Then
                Me.Master.SetDisplayMessage("Virtual Redirect INSERT Successful", Management.MessageType.ConfirmationMessage)
            Else
                Me.Master.SetDisplayMessage("Error INSERTING Virtual Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)
            End If

        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error INSERTING Virtual Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)

        End Try

    End Sub

    Protected Sub dsVirtualRedirect_Inserting(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceCommandEventArgs) Handles dsVirtualRedirect.Inserting
        Try
            dsVirtualRedirect.UpdateParameters("LastUpdatedBy").DefaultValue = Me.User.Identity.Name


        Catch ex As Exception

        End Try
    End Sub

    Protected Sub dsVirtualRedirect_Updated(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.SqlDataSourceStatusEventArgs) Handles dsVirtualRedirect.Updated
        Try
            'Check for errors and display UPDATE status
            If IsNothing(e.Exception) Then
                Me.Master.SetDisplayMessage("Virtual Redirect UPDATE Successful", Management.MessageType.ConfirmationMessage)
            Else
                Me.Master.SetDisplayMessage("Error UPDATING Virtual Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)
            End If

        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error UPDATING Virtual Redirect:  " & e.Exception.Message, Management.MessageType.ErrorMessage)

        End Try

    End Sub


    Protected Sub btnAddNew_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnAddNew.Click

        Try
            'Set Insert Paramaters
            dsVirtualRedirect.InsertParameters("Name").DefaultValue = txtName.Text
            dsVirtualRedirect.InsertParameters("RedirectURL").DefaultValue = txtURL.Text


            If IsNumeric(Page.User.Identity.Name) Then
                dsVirtualRedirect.InsertParameters("CreatedBy").DefaultValue = Page.User.Identity.Name
            Else
                dsVirtualRedirect.InsertParameters("CreatedBy").DefaultValue = 0
            End If



            dsVirtualRedirect.InsertParameters("Active").DefaultValue = chkActive.Checked

            'Insert Record
            dsVirtualRedirect.Insert()

            'Rebind the datagrid
            Me.gvVirtualRedirect.DataBind()

            'Clear existing form.
            Me.txtName.Text = ""
            Me.txtURL.Text = ""
            Me.chkActive.Checked = False


        Catch ex As Exception
            Me.Master.SetDisplayMessage("Error INSERTING Virtual Redirect:  " & ex.Message.ToString, Management.MessageType.ErrorMessage)
        End Try

    End Sub
End Class
