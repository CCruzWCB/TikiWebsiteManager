Imports Management

Partial Class MasterPages_NonSecure
    Inherits System.Web.UI.MasterPage

    Private _PageTitle As String

    Private _Title As String
    Public WriteOnly Property myTitle() As String
        Set(ByVal value As String)
            _Title = value
            Me.myPageTitle.Text = _Title
        End Set
    End Property

    Public Property PageTitle() As String
        Get
            Return _PageTitle
        End Get
        Set(ByVal value As String)
            _PageTitle = value
            Me.myPageTitle.Text = _PageTitle
        End Set
    End Property

    Protected Sub imgLogo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgLogo.Click
        Response.Redirect("~/default.aspx")
    End Sub

    Public Sub SetDisplayMessage(ByVal DisplayMessage As String, ByVal MessageType As MessageType)
        Me.lblInstruction.CssClass = ""
        Me.lblInstruction.Text = ""

        Select Case MessageType
            Case MessageType.ErrorMessage
                Me.lblInstruction.CssClass = "errorbox"
            Case MessageType.GeneralMessage
                Me.lblInstruction.CssClass = "general"
            Case MessageType.SyntaxMessage
                Me.lblInstruction.CssClass = "syntaxerror"
            Case MessageType.instructionMessage
                Me.lblInstruction.CssClass = "instructions"
            Case MessageType.WizardMessage
                Me.lblInstruction.CssClass = "wizard"
            Case MessageType.ConfirmationMessage
                Me.lblInstruction.CssClass = "confirmation"
        End Select

        Me.lblInstruction.Text = DisplayMessage

    End Sub
End Class

