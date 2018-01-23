
Imports System
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.IO
Imports System.Collections.Specialized
Imports System.ComponentModel
Imports System.Web.UI.Page

Namespace Management




    '*********************************************************************
    '
    ' HtmlTextBox Class
    '
    ' This control displays an HTML editor for uplevel browsers (ie 5.5 and higher)
    ' and a normal textarea for downlevel browsers.
    ' 
    ' Note: This control is based on the the MSDN Editor sample at:
    '
    '      http://msdn.microsoft.com/library/default.asp?url=/workshop/author/dhtml/reference/properties/canhavehtml.asp
    ' 
    ' Some of the client-side functionality of this control is based on
    ' Nikhil Kothari and Vandana Datye's HtmlEditor control 
    ' (read their book Developing Microsoft ASP.NET Server Controls and Components 
    '  -- it's the BEST book on control development)
    '
    ' Notice that the control uses an hidden input field to preserve content
    ' between postbacks. This is necessary since saving text in a TextArea tag
    ' causes the content to be automatically HTMLEncoded on postback.
    '
    '*********************************************************************
    <ValidationPropertyAttribute("Text")> _
    Public Class HtmlTextBox
        Inherits WebControl
        Implements INamingContainer, IPostBackDataHandler


        Private _allowHtml As AllowHtml = AllowHtml.None
        Private _isUplevel As Boolean = False
        Private _renderUplevel As Boolean = False
        Private _columns As Integer = 50
        Private _rows As Integer = 13


        Public Event TextChanged As EventHandler



        '*********************************************************************
        '
        ' AllowHtml Property
        '
        ' Determines whether a user can enter HTML content. Possible values are:
        '
        '  * None - No HTML
        '  * Limited - HTML can only be added through toolbar
        '  * Full - User can edit HTML source    
        '
        '*********************************************************************

        Public Property AllowHtml() As AllowHtml
            Get
                Return _allowHtml
            End Get
            Set(ByVal value As AllowHtml)
                _allowHtml = Value
            End Set
        End Property


        '*********************************************************************
        '
        ' IsUplevel Property
        '
        ' Enables user to force HtmlTextEditor into downlevel mode    
        '
        '*********************************************************************

        Public Property IsUplevel() As Boolean
            Get
                Return _isUplevel
            End Get
            Set(ByVal value As Boolean)
                _isUplevel = Value
            End Set
        End Property




        '*********************************************************************
        '
        ' Text Property
        '
        ' Represents the contents of the editor.
        ' Notice that we HTML encode in the special case of a downlevel
        ' browser with limited HTML enabled.
        '
        '*********************************************************************

        Public Property [Text]() As String
            Get
                If ViewState("Text") Is Nothing Then
                    Return String.Empty
                End If
                If _allowHtml = AllowHtml.Limited AndAlso _isUplevel = False Then
                    Return SimpleHtmlEncode(CStr(ViewState("Text")))
                End If

                Return CStr(ViewState("Text"))
            End Get
            Set(ByVal value As String)
                ViewState("Text") = Value
            End Set
        End Property




        '*********************************************************************
        '
        ' SimpleHtmlEncode Method
        '
        ' Replaces < with &lt; and > with &gt;    
        '
        '*********************************************************************
        Private Function SimpleHtmlEncode(ByVal [text] As String) As String
            [text] = [text].Replace("<", "&lt;")
            [text] = [text].Replace(">", "&gt;")
            Return [text]
        End Function 'SimpleHtmlEncode


        '*********************************************************************
        '
        ' Columns Property
        '
        ' Only used for downlevel browsers    
        '
        '*********************************************************************

        Public Property Columns() As Integer
            Get
                Return _columns
            End Get
            Set(ByVal value As Integer)
                _columns = Value
            End Set
        End Property

        '*********************************************************************
        '
        ' Rows Property
        '
        ' Only used for downlevel browsers    
        '
        '*********************************************************************

        Public Property Rows() As Integer
            Get
                Return _rows
            End Get
            Set(ByVal value As Integer)
                _rows = Value
            End Set
        End Property




        '*********************************************************************
        '
        ' DetermineUplevel Method
        '
        ' Only render client script for IE 5.5 or higher    
        '
        '*********************************************************************
        Protected Sub DetermineUplevel()
            If Not (Context Is Nothing) Then
                Dim _browser As HttpBrowserCapabilities = Context.Request.Browser

                Dim _hasEcmaScript As Boolean = _browser.EcmaScriptVersion.CompareTo(New Version(1, 2)) >= 0
                Dim _hasDom As Boolean = _browser.MSDomVersion.Major >= 4
                Dim _hasBehaviors As Boolean = _browser.MajorVersion > 5 OrElse (_browser.MajorVersion = 5 AndAlso _browser.MinorVersion >= 0.5)
                _isUplevel = _hasEcmaScript AndAlso _hasDom AndAlso _hasBehaviors
            End If
        End Sub 'DetermineUplevel





        '*********************************************************************
        '
        ' LoadPostData Property
        '
        ' Updates control properties when page is posted
        '
        '*********************************************************************
        Public Function LoadPostData(ByVal postDataKey As String, ByVal values As NameValueCollection) As Boolean Implements IPostBackDataHandler.LoadPostData
            If Not [Text].Equals(values(UniqueID)) Then
                [Text] = values(UniqueID)
                Return True
            End If
            Return False
        End Function 'LoadPostData



        '*********************************************************************
        '
        ' RaisPostDataChangedEvent Method
        '
        ' Raises the TextChanged event
        '
        '*********************************************************************            
        Public Overridable Sub RaisePostDataChangedEvent() Implements IPostBackDataHandler.RaisePostDataChangedEvent
            OnTextChanged(EventArgs.Empty)
        End Sub 'RaisePostDataChangedEvent


        Protected Overridable Sub OnTextChanged(ByVal e As EventArgs)
            RaiseEvent TextChanged(Me, e)
        End Sub 'OnTextChanged




        '*********************************************************************
        '
        ' GetClientIncludes Method
        '
        ' Returns the client scripts    
        '
        '*********************************************************************
        Private Function GetClientIncludes() As String
            Return String.Format("<script language=""JavaScript"" src=""{0}""></script>" + vbCr + vbLf + "<?xml:namespace prefix=""community""/>" + vbCr + vbLf + "<?import namespace=""community"" implementation=""{1}""/>", Page.ResolveUrl("~/js/HtmlTextBox.js"), Page.ResolveUrl("~/HtmlControls/HtmlTextBox.htc"))
        End Function 'GetClientIncludes



        '*********************************************************************
        '
        ' EmoticonPath Property
        '
        ' Helper property for retrieving path of folder with emoticon icons    
        '
        '*********************************************************************

        ReadOnly Property EmoticonPath() As String
            Get
                Return ResolveUrl("CBWebsiteManager/Images/Emoticons/")
            End Get
        End Property


        '*********************************************************************
        '
        ' EmoticonPath Property
        '
        ' Helper property for retrieving path of folder with emoticon icons    
        '
        '*********************************************************************

        ReadOnly Property InsertImagePath() As String
            Get
                Return ResolveUrl("CBWebsiteManager/Images/Inserts/")
            End Get
        End Property

        '*********************************************************************
        '
        ' EmoticonList Property
        '
        ' Returns comma delimited list of emoticon images    
        '
        '*********************************************************************

        ReadOnly Property EmoticonList() As String
            Get
                Dim strList As String = String.Empty
                Try
                    Dim arrFiles As String() = Directory.GetFiles(Page.MapPath(EmoticonPath), "*.gif")
                    Dim i As Integer
                    For i = 0 To arrFiles.Length - 1
                        If i = 0 Then
                            strList = "'" + Path.GetFileName(arrFiles(i)) + "'"
                        Else
                            strList += ",'" + Path.GetFileName(arrFiles(i)) + "'"
                        End If
                    Next i
                Catch ex As Exception

                End Try

                Return strList
            End Get
        End Property

        '*********************************************************************
        '
        ' InsertImageList Property
        '
        ' Returns comma delimited list of images from specified folder.   
        '
        '*********************************************************************

        ReadOnly Property InsertImageList() As String
            Get
                Dim strList As String = String.Empty
                Try
                    Dim arrFiles As String() = Directory.GetFiles(Page.MapPath(InsertImagePath), "*.gif")
                    Dim i As Integer
                    For i = 0 To arrFiles.Length - 1
                        If i = 0 Then
                            strList = "'" + Path.GetFileName(arrFiles(i)) + "'"
                        Else
                            strList += ",'" + Path.GetFileName(arrFiles(i)) + "'"
                        End If
                    Next i
                Catch ex As Exception

                End Try

                Return strList
            End Get
        End Property






        '*********************************************************************
        '
        ' OnPreRender Method
        '
        ' If browser is uplevel, add the client scripts to the page    
        '
        '*********************************************************************
        Protected Overrides Sub OnPreRender(ByVal e As EventArgs)
            MyBase.OnPreRender(e)

            ' Determine render uplevel
            _renderUplevel = _isUplevel AndAlso _allowHtml <> AllowHtml.None



            If _renderUplevel Then
                Dim scriptKey As String = GetType(HtmlTextBox).FullName

                ' Register the editor in list of editors
                Page.ClientScript.RegisterArrayDeclaration("htmlDesignerList", "'" + ClientID + "'")


                ' add behavior namespace and emoticonlist
                If Not Page.ClientScript.IsClientScriptBlockRegistered(scriptKey) Then
                    Page.ClientScript.RegisterClientScriptBlock(GetType(HtmlTextBox), scriptKey, GetClientIncludes())
                    Page.ClientScript.RegisterArrayDeclaration("emoticonList", EmoticonList)
                    Page.ClientScript.RegisterArrayDeclaration("InsertImageList", InsertImageList)
                End If

                ' Register startup script (only once for all instances)
                If Not Page.ClientScript.IsStartupScriptRegistered(scriptKey) Then
                    Page.ClientScript.RegisterStartupScript(GetType(HtmlTextBox), scriptKey, String.Format("<script language=""JavaScript"">var appBasePath='{0}'; htb_InitializeElements()</script>", ConfigurationManager.AppSettings("WebSiteAdminHome")))
                End If
            End If
        End Sub 'OnPreRender




        '*********************************************************************
        '
        ' Render Method
        '
        ' If downlevel (or allowHtml is none) render a textarea, otherwise
        ' render an HTML editor    
        '
        '*********************************************************************
        Protected Overrides Sub Render(ByVal writer As HtmlTextWriter)
            If _renderUplevel Then
                RenderUplevel(writer)
            Else
                RenderDownlevel(writer)
            End If
        End Sub 'Render



        '*********************************************************************
        '
        ' RenderDownlevel Method
        '
        ' For downlevel browsers, we just render a textarea.
        '
        '*********************************************************************
        Private Sub RenderDownlevel(ByVal writer As HtmlTextWriter)
            writer.RenderBeginTag(HtmlTextWriterTag.Span)
            writer.AddAttribute(HtmlTextWriterAttribute.Id, ClientID)
            writer.AddAttribute(HtmlTextWriterAttribute.Name, UniqueID)
            writer.AddAttribute(HtmlTextWriterAttribute.Cols, _columns.ToString())
            writer.AddAttribute(HtmlTextWriterAttribute.Rows, _rows.ToString())
            writer.RenderBeginTag(HtmlTextWriterTag.Textarea)
            writer.Write([Text])
            writer.RenderEndTag()
        End Sub 'RenderDownlevel




        '*********************************************************************
        '
        ' RenderDownlevel Method
        '
        ' For downlevel browsers, we just render a textarea.
        '
        '*********************************************************************
        Private Sub RenderUplevel(ByVal writer As HtmlTextWriter)
            writer.RenderBeginTag(HtmlTextWriterTag.Span)
            writer.AddAttribute(HtmlTextWriterAttribute.Id, ClientID)
            writer.AddAttribute(HtmlTextWriterAttribute.Name, UniqueID)
            writer.AddAttribute(HtmlTextWriterAttribute.Type, "hidden")
            writer.AddAttribute(HtmlTextWriterAttribute.Value, [Text])
            writer.RenderBeginTag(HtmlTextWriterTag.Input)
            writer.RenderEndTag()

            writer.WriteLine()

            ' Next, do the HTML Designer
            If ControlStyleCreated Then
                ControlStyle.AddAttributesToRender(writer, Me)
            End If

            writer.AddAttribute(HtmlTextWriterAttribute.Id, ClientID + "HtmlDesigner")
            writer.AddAttribute("allowhtml", _allowHtml.ToString())
            writer.AddAttribute("onHtmlChanged", "htb_OnHtmlChanged(this, document.all['" + ClientID + "'])", False)
            writer.RenderBeginTag("community:HtmlDesigner")
            writer.RenderEndTag()

            writer.RenderEndTag() 'span        
        End Sub 'RenderUplevel





        '*********************************************************************
        '
        ' HtmlTextBox Constructor
        '
        ' Set the base TextBox control to MultiLine mode
        ' and determine AllowHTML from context    
        '
        '*********************************************************************
        Public Sub New()
            Width = New Unit("500px")
            Height = New Unit("200px")
            DetermineUplevel()
            If Not (Context Is Nothing) AndAlso Not (Context.Items("SectionInfo") Is Nothing) Then
                _allowHtml = AllowHtml.Full
            End If
        End Sub 'New 
    End Class 'HtmlTextBox 



    '*********************************************************************
    '
    ' AllowHtml Enumeration
    '
    ' Possible values for AllowHtml    
    '
    '*********************************************************************

    Public Enum AllowHtml
        None = 0
        Limited = 1
        Full = 2
    End Enum 'AllowHtml


End Namespace