
Partial Class WebControls_Calendar
    Inherits System.Web.UI.Page

    Protected Sub Calendar_DayRender(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DayRenderEventArgs) Handles Calendar.DayRender
        If e.Day.Date = DateTime.Now.ToString("d") Then
            e.Cell.BackColor = System.Drawing.Color.LightGray
        End If

    End Sub

    Protected Sub Calendar_SelectionChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Calendar.SelectionChanged
        Dim strjscript As String = "<script language=""javascript"">"
        strjscript &= "window.opener." & _
              HttpContext.Current.Request.QueryString("formname") & ".value = '" & _
              Calendar.SelectedDate & "';window.close();"
        strjscript = strjscript & "</script" & ">" 'Don't Ask, Tool Bug

        lblDate.Text = strjscript  'Set the literal control's text to the JScript code

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

    End Sub
End Class
