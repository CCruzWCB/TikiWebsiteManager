Imports System
Imports System.Web
Imports System.Data.SqlClient
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Web.Configuration
Imports System.Collections.Generic
Imports System.Runtime.CompilerServices
Imports System.Configuration.Provider
Imports System.Security.Permissions
Imports System.Data.Common
Imports System.Data

'"SqlSiteMapProvider" Class Profile ******************************************************
'
'       Class: SqlSiteMapProvider 
'       Parent File: SQLSitemapProvider.vb
'       Parent Project: GLWebsite
'       Parent Solution: GLSolution
'       Created By: Amy Cook
'       Created on: 04/27/07
'       Description:  This class is the provider for the Site Navigation, including the
'                     menu control as well as the SiteMapPath or "breadcrumb" control.
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


<SqlClientPermission(SecurityAction.Demand, Unrestricted:=True)> _
Public Class SqlSiteMapProvider1
    Inherits StaticSiteMapProvider
    Private Const _errmsg1 As String = "Missing node ID"
    Private Const _errmsg2 As String = "Duplicate node ID"
    Private Const _errmsg3 As String = "Missing parent ID"
    Private Const _errmsg4 As String = "Invalid parent ID"
    Private Const _errmsg5 As String = "Empty or missing connectionStringName"
    Private Const _errmsg6 As String = "Missing connection string"
    Private Const _errmsg7 As String = "Empty connection string"


    Private _connect As String
    Private _indexID, _indexTitle, _indexUrl, _indexDesc, _indexRoles, _indexParent As Integer
    Private _nodes As Dictionary(Of Integer, SiteMapNode) = New Dictionary(Of Integer, SiteMapNode)(16)
    Private _root As SiteMapNode

    Public Overrides Sub Initialize(ByVal name As String, _
                                ByVal config As NameValueCollection)
        ' Verify that config isn't null\\


        _root = Nothing ' temp code to ensure we always run this code during testing
        _nodes = New Dictionary(Of Integer, SiteMapNode)(16)
        Clear()


        If config Is Nothing Then
            Throw New ArgumentNullException("config")
        End If

        ' Assign the provider a default name if it doesn't have one
        If String.IsNullOrEmpty(name) Then
            name = "SqlSiteMapProvider"
        End If

        ' Add a default "description" attribute to config if the
        ' attribute doesn't exist or is empty
        If String.IsNullOrEmpty(config("description")) Then
            config.Remove("description")
            config.Add("description", "SQL site map provider")
        End If

        ' Call the base class's Initialize method
        MyBase.Initialize(name, config)

        ' Initialize _connect
        Dim connect As String = config("connectionStringName")

        If String.IsNullOrEmpty(connect) Then
            Throw New ProviderException(_errmsg5)
        End If

        config.Remove("connectionStringName")

        If WebConfigurationManager.ConnectionStrings(connect) _
            Is Nothing Then
            Throw New ProviderException(_errmsg6)
        End If

        _connect = _
        WebConfigurationManager.ConnectionStrings( _
                                        connect).ConnectionString

        If String.IsNullOrEmpty(_connect) Then
            Throw New ProviderException(_errmsg7)
        End If

        ' In beta 2, SiteMapProvider processes the
        ' securityTrimmingEnabled attribute but fails to remove it.
        ' Remove it now so we can check for unrecognized
        ' configuration attributes.

        'If Not config("securityTrimmingEnabled") Is Nothing Then
        '    config.Remove("securityTrimmingEnabled")
        'End If

        ' Throw an exception if unrecognized attributes remain
        If config.Count > 0 Then
            Dim attr As String = config.GetKey(0)
            If (Not String.IsNullOrEmpty(attr)) Then
                Throw New  _
                ProviderException("Unrecognized attribute: " & attr)
            End If
        End If
    End Sub

    Public Overrides Function BuildSiteMap() As SiteMapNode
        SyncLock Me
            ' Return immediately if this method has been called before

            If Not _root Is Nothing Then
                Return _root
            End If

            ' Query the database for site map nodes
            Dim connection As SqlConnection = New SqlConnection(_connect)

            Try
                connection.Open()
                'Dim command As SqlCommand = New SqlCommand("SELECT ProductCategoryID, Name, Description, nullif(ParentCategoryID,0) as ParentCategoryID ,NULL AS Roles, '~/Ecommerce/Category.aspx?CategoryID=' + convert(varchar(5),ProductCategoryID) as url FROM vProductCategories where CompanyID = " & ConfigurationManager.AppSettings("CompanyID") & " and BrandID = " & ConfigurationManager.AppSettings("BrandID") & " and active = 1 order by parentcategoryID", connection)
                'Dim command As SqlCommand = New SqlCommand("SELECT ProductCategoryID, Name, Description, nullif(ParentCategoryID,-1) as ParentCategoryID ,NULL AS Roles, '~/Ecommerce/Category.aspx?CategoryID=' + convert(varchar(5),ProductCategoryID) as url FROM vProductCategories where CompanyID = " & ConfigurationManager.AppSettings("CompanyID") & " and BrandID = " & ConfigurationManager.AppSettings("BrandID") & " and active = 1 order by parentcategoryID", connection)
                'command.CommandType = CommandType.Text
                Dim command As New SqlCommand("[GetEcommerceManageMenuByBrand]", connection)


                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ConfigurationManager.AppSettings("BrandID")))

                Dim reader As SqlDataReader = command.ExecuteReader()
                _indexID = reader.GetOrdinal("ProductCategoryID")
                _indexUrl = reader.GetOrdinal("Url")
                _indexTitle = reader.GetOrdinal("Name")
                _indexDesc = reader.GetOrdinal("Description")
                _indexRoles = reader.GetOrdinal("Roles")
                _indexParent = reader.GetOrdinal("ParentCategoryID")


                ' Create the root SiteMapNode and add it to
                ' the site map

                ' Build a tree of SiteMapNodes underneath
                ' the root node
                Do While reader.Read()

                    If reader.IsDBNull(_indexParent) Then
                        _root = CreateSiteMapNodeFromDataReader(reader)
                        AddNode(_root, Nothing)
                    Else
                        ' Create another site map node and
                        ' add it to the site map
                        Dim node As SiteMapNode = _
                            CreateSiteMapNodeFromDataReader(reader)
                        AddNode(node, _
                                GetParentNodeFromDataReader(reader))

                    End If

                Loop

            Finally
                connection.Close()
            End Try

            ' Return the root SiteMapNode
            Return _root

        End SyncLock
    End Function

    Protected Overrides Function GetRootNodeCore() As SiteMapNode
        BuildSiteMap()
        Return _root
    End Function

    ' Helper methods
    Private Function CreateSiteMapNodeFromDataReader( _
                        ByVal reader As DbDataReader) As SiteMapNode
        ' Make sure the node ID is present
        If reader.IsDBNull(_indexID) Then
            Throw New ProviderException(_errmsg1)
        End If

        ' Get the node ID from the DataReader
        Dim id As Integer = reader.GetInt32(_indexID)

        ' Make sure the node ID is unique
        If _nodes.ContainsKey(id) Then
            Throw New ProviderException(_errmsg2)
        End If

        ' Get title, URL, description, and roles from the DataReader
        Dim title As String
        If reader.IsDBNull(_indexTitle) Then
            title = Nothing
        Else
            title = LCase(reader.GetString(_indexTitle).Trim())
        End If
        Dim url As String
        If reader.IsDBNull(_indexUrl) Then
            url = Nothing
        Else
            url = reader.GetString(_indexUrl).Trim()
        End If
        Dim description As String
        If reader.IsDBNull(_indexDesc) Then
            description = Nothing
        Else
            description = reader.GetString(_indexDesc).Trim()
        End If
        Dim roles As String
        If reader.IsDBNull(_indexRoles) Then
            roles = Nothing
        Else
            roles = reader.GetString(_indexRoles).Trim()
        End If

        ' If roles were specified, turn the list into a string array
        Dim rolelist As String() = Nothing
        If (Not String.IsNullOrEmpty(roles)) Then
            rolelist = roles.Split(New Char() {","c, ";"c}, 512)
        End If

        ' Create a SiteMapNode
        Dim node As SiteMapNode = _
            New SiteMapNode(Me, id.ToString(), url, title, _
            description, rolelist, Nothing, Nothing, Nothing)

        ' Record the node in the _nodes dictionary
        _nodes.Add(id, node)

        ' Return the node        
        Return node
    End Function

    Private Function GetParentNodeFromDataReader( _
                        ByVal reader As DbDataReader) As SiteMapNode
        ' Make sure the parent ID is present
        If reader.IsDBNull(_indexParent) Then
            Throw New ProviderException(_errmsg3)
        End If

        ' Get the parent ID from the DataReader
        Dim pid As Integer = reader.GetInt32(_indexParent)

        ' Make sure the parent ID is valid
        If (Not _nodes.ContainsKey(pid)) Then
            Throw New ProviderException(_errmsg4)
        End If

        ' Return the parent SiteMapNode
        Return _nodes(pid)
    End Function

    Public Overrides ReadOnly Property CurrentNode() As SiteMapNode

        Get

            Dim currentUrl As String = FindCurrentUrl()

            ' Build the site map in memory.

            BuildSiteMap()

            ' Find the SiteMapNode that represents the current page.

            Dim aCurrentNode As SiteMapNode = FindSiteMapNode(currentUrl)

            Return aCurrentNode

        End Get

    End Property
    Private Function FindCurrentUrl() As String

        ' The current HttpContext.

        Dim currentContext As HttpContext = HttpContext.Current

        If Not (currentContext Is Nothing) Then

            Return currentContext.Request.RawUrl

        Else

            Return currentContext.Request.Path

        End If

    End Function 'FindCurrentUrl

    Public Overrides Function FindSiteMapNode(ByVal rawUrl As String) As SiteMapNode

        ' Does the root node match the URL?

        If RootNode.Url = rawUrl Then

            Return RootNode

        Else

            Dim candidate As SiteMapNode = Nothing

            ' Retrieve the SiteMapNode that matches the URL.

            SyncLock Me

                candidate = GetNode(_root, rawUrl)

            End SyncLock

            Return candidate

        End If

    End Function 'FindSiteMapNode

    Private Function GetNode(ByVal nodes As SiteMapNode, ByVal url As String) As SiteMapNode

        Dim thisNode As SiteMapNode
        Dim childNode As SiteMapNode

        For Each thisNode In nodes.ChildNodes

            If CStr(thisNode.Url) = url Or CStr(thisNode.Url) = Replace(url, "/GLWebsite", "~") Then
                Return thisNode
            Else
                For Each childNode In thisNode.ChildNodes
                    If CStr(childNode.Url) = url Or CStr(childNode.Url) = Replace(url, "/GLWebsite", "~") Then
                        Return childNode
                    End If
                Next
            End If

        Next

        Return Nothing

    End Function 'GetNode


End Class










