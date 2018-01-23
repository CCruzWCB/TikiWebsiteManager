Imports System
Imports System.Web
Imports System.Data.SqlClient
Imports System.Collections.Specialized
Imports System.Configuration
Imports System.Web.Configuration
Imports System.Collections.Generic
Imports System.Configuration.Provider
Imports System.Security.Permissions
Imports System.Data.Common
Imports System.Data
Imports System.Web.Caching




<SqlClientPermission(SecurityAction.Demand, Unrestricted:=True)> _
Public Class SqlSiteMapProvider
    Inherits StaticSiteMapProvider
    Private Const _errmsg1 As String = "Missing node ID"
    Private Const _errmsg2 As String = "Duplicate node ID"
    Private Const _errmsg3 As String = "Missing parent ID"
    Private Const _errmsg4 As String = "Invalid parent ID"
    Private Const _errmsg5 As String = "Empty or missing connectionStringName"
    Private Const _errmsg6 As String = "Missing connection string"
    Private Const _errmsg7 As String = "Empty connection string"
    Private Const _errmsg8 As String = "Invalid sqlCacheDependency"
    Private Const _cacheDependencyName As String = "__SiteMapCacheDependency"

    Private _connect As String
    Private _database As String
    Private _table As String
    Private _2005dependency As Boolean = False
    Private _indexID As Integer
    Private _indexTitle As Integer
    Private _indexUrl As Integer
    Private _indexDesc As Integer
    Private _indexRoles As Integer
    Private _indexParent As Integer
    Private _nodes As Dictionary(Of Integer, SiteMapNode) = New Dictionary(Of Integer, SiteMapNode)(16)
    Private ReadOnly _lock As Object = New Object
    Private _root As SiteMapNode

    'Added...Declare an arraylist to hold all the roles this menu item applies to
    Public roles As New ArrayList

    Public Overloads Overrides Sub Initialize(ByVal name As String, ByVal config As NameValueCollection)
        ' Verify that config isn't null
        If config Is Nothing Then
            Throw New ArgumentNullException("config")
        End If

        ' Assign the provider a default name if it doesn't have one
        If String.IsNullOrEmpty(name) Then
            name = "SqlSiteMapProvider"
        End If

        ' Add a default "description" attribute to config if the
        ' attribute doesn’t exist or is empty
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

        If WebConfigurationManager.ConnectionStrings(connect) Is Nothing Then
            Throw New ProviderException(_errmsg6)
        End If

        _connect = WebConfigurationManager.ConnectionStrings(connect).ConnectionString

        If String.IsNullOrEmpty(_connect) Then
            Throw New ProviderException(_errmsg7)
        End If


        ' Initialize SQL cache dependency info
        Dim dependency As String = config("sqlCacheDependency")

        If Not [String].IsNullOrEmpty(dependency) Then
            If [String].Equals(dependency, "CommandNotification", StringComparison.InvariantCultureIgnoreCase) Then
                SqlDependency.Start(_connect)
                _2005dependency = True
            Else
                ' If not "CommandNotification", then extract database and table names
                Dim info As String() = dependency.Split(New Char() {":"c})
                If info.Length <> 2 Then
                    Throw New ProviderException(_errmsg8)
                End If
                _database = info(0)
                _table = info(1)
            End If

            config.Remove("sqlCacheDependency")

        End If


        ' SiteMapProvider processes the securityTrimmingEnabled
        ' attribute but fails to remove it. Remove it now so we can
        ' check for unrecognized configuration attributes.
        If Not (config("securityTrimmingEnabled") Is Nothing) Then
            config.Remove("securityTrimmingEnabled")
        End If

        ' Throw an exception if unrecognized attributes remain
        If config.Count > 0 Then
            Dim attr As String = config.GetKey(0)
            If Not [String].IsNullOrEmpty(attr) Then
                Throw New ProviderException("Unrecognized attribute: " + attr)
            End If
        End If


    End Sub

    Public Overloads Overrides Function BuildSiteMap() As SiteMapNode

        SyncLock _lock

            ' Return immediately if this method has been called before

            If _root IsNot Nothing Then

                Return _root

            End If



            ' Query the database for site map nodes

            Dim connection As New SqlConnection(_connect)



            Try

                'Dim command As New SqlCommand("GetMarketingMenu", connection)
                Dim command As New SqlCommand("LLFBPS.dbo.[GetEcommerceManageMenuByBrand]", connection)


                command.CommandType = CommandType.StoredProcedure
                command.Parameters.Add(New System.Data.SqlClient.SqlParameter("@BrandID", System.Data.SqlDbType.Int, 4, ParameterDirection.Input, False, CType(0, Byte), CType(0, Byte), "", DataRowVersion.Current, ConfigurationManager.AppSettings("BrandID")))



                'SqlCacheDependencyAdmin.EnableNotifications(_connect)
                'SqlCacheDependencyAdmin.EnableTableForNotifications(_connect, "LLFBPS.dbo.Product")
                'Dim Tables() As String = SqlCacheDependencyAdmin.GetTablesEnabledForNotifications(_connect)







                'Why create a dependency for ProductAttributes

                ' Create a SQL cache dependency if requested

                Dim dependency As SqlCacheDependency = Nothing



                If _2005dependency Then

                    dependency = New SqlCacheDependency(command)

                ElseIf Not [String].IsNullOrEmpty(_database) AndAlso Not String.IsNullOrEmpty(_table) Then

                    dependency = New SqlCacheDependency(_database, _table)
                Else
                    dependency = New SqlCacheDependency("LLFBPS", "Product")


                End If



                connection.Open()

                Dim reader As SqlDataReader = command.ExecuteReader()

                _indexID = reader.GetOrdinal("ProductCategoryID")

                _indexUrl = reader.GetOrdinal("Url")

                _indexTitle = reader.GetOrdinal("Name")

                _indexDesc = reader.GetOrdinal("Description")

                _indexRoles = reader.GetOrdinal("Roles")

                _indexParent = reader.GetOrdinal("ParentCategoryID")



                If reader.Read() Then

                    ' Create the root SiteMapNode and add it to the site map

                    _root = CreateSiteMapNodeFromDataReader(reader)

                    AddNode(_root, Nothing)



                    ' Build a tree of SiteMapNodes underneath the root node

                    While reader.Read()

                        ' Create another site map node and add it to the site map

                        Dim node As SiteMapNode = CreateSiteMapNodeFromDataReader(reader)

                        AddNode(node, GetParentNodeFromDataReader(reader))
                        Dim CategoryName As String = reader(1)

                    End While



                    ' Use the SQL cache dependency

                    If dependency IsNot Nothing Then

                        'HttpRuntime.Cache.Insert(_cacheDependencyName, New Object(), dependency, Cache.NoAbsoluteExpiration, Cache.NoSlidingExpiration, CacheItemPriority.NotRemovable, New CacheItemRemovedCallback(AddressOf OnSiteMapChanged))
                        HttpRuntime.Cache.Insert(_cacheDependencyName, New Object(), dependency, DateTime.Now.AddSeconds(30), Cache.NoSlidingExpiration, CacheItemPriority.Normal, New CacheItemRemovedCallback(AddressOf OnSiteMapChanged))

                    End If

                End If

            Finally

                connection.Close()

            End Try



            ' Return the root SiteMapNode

            Return _root

        End SyncLock

    End Function



    Protected Overloads Overrides Function GetRootNodeCore() As SiteMapNode

        SyncLock _lock

            BuildSiteMap()

            Return _root

        End SyncLock

    End Function



    ' Helper methods

    Private Function CreateSiteMapNodeFromDataReader(ByVal reader As DbDataReader) As SiteMapNode

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

        Dim title As String = ReplaceNullRefs(reader, _indexTitle)

        Dim url As String = ReplaceNullRefs(reader, _indexUrl)



        'Eliminated...see http://weblogs.asp.net/psteele/archive/2003/10/09/31250.aspx

        Dim description As String = ReplaceNullRefs(reader, _indexDesc)



        'Changed variable name from 'roles' to 'rolesN' and added line 230 to dump all roles into an arrayList

        'Dim rolesN As String = IIf(reader.IsDBNull(_indexRoles), Nothing, reader.GetString(_indexRoles).Trim())





        'Dim rolelist As String() = Nothing

        'If Not [String].IsNullOrEmpty(rolesN) Then

        '    rolelist = rolesN.Split(New Char() {","c, ";"c}, 512)

        'End If

        'roles = ArrayList.Adapter(rolelist)
        Dim metakeywords As String = title
        Dim metadescription As String = description
        Dim CategoryID As String = id


        Dim attributeString As String = "categoryid:" & CategoryID & ";keywords:" & metakeywords & ";desc:" & metadescription & ";"
        Dim atts As New NameValueCollection()
        If Not String.IsNullOrEmpty(attributeString) Then


            atts.Add("keywords", metakeywords)
            atts.Add("desc", metadescription)
            atts.Add("categoryid", CategoryID)

        End If





        ' Create a SiteMapNode

        Dim node As New SiteMapNode(Me, id.ToString(), url, title, description, Nothing, atts, Nothing, Nothing)




        ' Record the node in the _nodes dictionary

        _nodes.Add(id, node)



        ' Return the node        

        Return node

    End Function



    Private Function ReplaceNullRefs(ByVal rdr As DbDataReader, ByVal rdrVal As Integer) As String

        If Not (rdr.IsDBNull(rdrVal)) Then



            ' Thanks Rob Johnston 

            Return rdr.GetString(rdrVal)

        Else

            Return String.Empty

        End If

    End Function



    Private Function GetParentNodeFromDataReader(ByVal reader As DbDataReader) As SiteMapNode

        ' Make sure the parent ID is present

        If reader.IsDBNull(_indexParent) Then

            '**** Commented out throw, added exit function ****

            Throw New ProviderException(_errmsg3)

            'Exit Function

        End If



        ' Get the parent ID from the DataReader

        Dim pid As Integer = reader.GetInt32(_indexParent)



        ' Make sure the parent ID is valid

        If Not _nodes.ContainsKey(pid) Then

            Throw New ProviderException(_errmsg4)

        End If



        ' Return the parent SiteMapNode

        Return _nodes(pid)

    End Function



    Private Sub OnSiteMapChanged(ByVal key As String, ByVal item As Object, ByVal reason As CacheItemRemovedReason)

        SyncLock _lock

            If key = _cacheDependencyName AndAlso (reason = CacheItemRemovedReason.DependencyChanged OrElse reason = CacheItemRemovedReason.Expired) Then

                ' Refresh the site map

                Clear()

                _nodes.Clear()

                _root = Nothing

            End If

        End SyncLock

    End Sub

End Class