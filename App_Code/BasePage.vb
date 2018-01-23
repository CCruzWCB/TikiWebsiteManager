Imports Microsoft.VisualBasic
Imports System
Imports System.Data
Imports System.Configuration
Imports System.Web
Imports System.Web.Security
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Web.UI.WebControls.WebParts
Imports System.Web.UI.HtmlControls
Imports Management.Security
Public MustInherit Class BasePage
    Inherits System.Web.UI.Page

    Private _runtimeMasterPageFile As String
    Private _SectionType As ResourceSectionType

   
    Public Property SectionType() As ResourceSectionType
        Get
            Return _SectionType
        End Get
        Set(ByVal value As ResourceSectionType)
            _SectionType = value
        End Set
    End Property

    Public Property RuntimeMasterPageFile() As String
        Get
            Return _runtimeMasterPageFile
        End Get
        Set(ByVal value As String)
            _runtimeMasterPageFile = value
        End Set

    End Property


    Protected Overrides Sub OnPreInit(ByVal e As System.EventArgs)
        If Not Page.IsPostBack Then
            If Not String.IsNullOrEmpty(_SectionType) Then
                'Check to see if user has access to the requested function 

                Dim strCombineLists As String = ""
                Dim IsValid As Boolean = False
                Dim ValidAdminUserID As Integer


                Dim ValidAdminUserIDList As String()

                If _SectionType > 0 Then


                    Select Case _SectionType
                        Case ResourceSectionType.RecipeManagement
                            'Can belong to the Recipe Admin section OR Site Admins
                            strCombineLists = ConfigurationManager.AppSettings("Application_RecipeAdminUserIDs").ToString & "," & ConfigurationManager.AppSettings("Application_AdminUserIDs").ToString
                            ValidAdminUserIDList = strCombineLists.Split(",")
                        Case Else
                            'MUST belong to the Site Admins
                            ValidAdminUserIDList = ConfigurationManager.AppSettings("Application_AdminUserIDs").ToString.Split(",")
                    End Select

                    Try
                        For Each ValidAdminUserID In ValidAdminUserIDList
                            If User.Identity.Name = ValidAdminUserID Then
                                IsValid = True
                                Exit For
                            End If
                        Next
                    Catch ex As Exception

                    End Try

                End If


                If Not IsValid Then
                    Response.Redirect(String.Format("{0}/default.aspx?message={1}&MessageTypeID={2}", ConfigurationManager.AppSettings("WebSiteAdminHome").ToString, "ACCESS DENIED: " & Page.Title, 3))
                End If



            End If
        End If
       


        MyBase.OnPreInit(e)
        'End If

    End Sub

   
    '' This Function is used to capture any Asp.NET generated error log and then send an email. It then 
    ''display a General Error Message or redirects to the error page.
    'Private Sub Page_Error(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Error

    '    Dim ErrorDetails As String = ""
    '    Dim Cxt As HttpContext
    '    Dim Ex As Exception

    '    Cxt = HttpContext.Current
    '    Ex = Server.GetLastError

    '    ErrorDetails &= "Page Source: " & Cxt.Request.Url.ToString & "<BR>"
    '    ErrorDetails &= "Source: " & Ex.Source.ToString & "<BR>"
    '    ErrorDetails &= "Message: " & Ex.Message.ToString & "<BR>"
    '    ErrorDetails &= "User ID: " & Cxt.User.Identity.Name & "<BR>"
    '    ErrorDetails &= "Stack Trace:" & Ex.StackTrace.ToString & "<BR>"

    '    'Log the error and email 
    '    'ApplicationLog.AddLogEntry(LogResourceType.System, 0, "CMS Application Error", "Details: <BR>" & ErrorDetails, HttpContext.Current.User.Identity.Name, SeverityLevel.High)

    '    'Clear the error
    '    Server.ClearError()

    '    'Write this message to the page 
    '    Response.Write("An Error Occurred processing this request:<BR> " & ErrorDetails)
    '    'Response.Redirect(String.Format("{0}CustomErrors/Error.aspx?aspxerrorpath={1}", ConfigurationManager.AppSettings("ApplicationHTTPRoot"), Cxt.Request.Path.ToString))






    'End Sub
End Class





