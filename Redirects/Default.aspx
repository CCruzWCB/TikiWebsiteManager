<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="Redirects_Default" title="Untitled Page" %>
<%@ MasterType  VirtualPath  ="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
<asp:SiteMapDataSource ID="SiteMapDataSourceSupport" runat="server" StartFromCurrentNode ="true"  
                                SiteMapProvider="DefaultSiteMapProvider" ShowStartingNode="false"></asp:SiteMapDataSource>
                                
                                <div style ="text-align:left;">
                                <asp:Repeater runat="server" ID="Repeater1" DataSourceID="SiteMapDataSourceSupport"
                                Visible="true">
                                <HeaderTemplate>
                                    <ul>
                                </HeaderTemplate>
                                <ItemTemplate>
                                    <li><a id="A1" runat="server" title='<%#Eval("description") %>' href='<%# Eval("Url") %>'>
                                        <span>
                                            <%#Eval("title")%></span></a>
                                        
                                    </li>
                                </ItemTemplate>
                                <FooterTemplate>
                                    </ul>
                                </FooterTemplate>
                            </asp:Repeater>
                            </div>
                                
</asp:Content>

