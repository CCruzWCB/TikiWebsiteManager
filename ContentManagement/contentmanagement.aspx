<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="contentmanagement.aspx.vb" Inherits="ContentManagement_contentmanagement" %>
<%@ MasterType  VirtualPath  ="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
<ul>
                           <li class="listitem"><a href="managehomepage.aspx">Home Page</a></li>
                                <li class="listitem"><a href="ManageSeriesDataCubes.aspx">Series Pages</a></li>
                                <li class="listitem"><a href="ManagePromotionDataCubes.aspx">Promotion Pages</a></li>
                                <li class="listitem"><a href="managecategorypage.aspx">Category Pages</a></li>
                          
                                </ul> 
                                
</asp:Content>
