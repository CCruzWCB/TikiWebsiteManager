<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false"
    CodeFile="Default.aspx.vb" Inherits="_Default" %>

<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ContentPlaceHolderID="contentMain" ID="ContentMain" runat="server">
    <table cellspacing="0" cellpadding="5" width="100%" border="0">
        <tr>
            <!-- Begin Manage Reports -->
            <td valign="top" style="width: 50%;">
                <table class="box" cellspacing="0" border="0" cellpadding="0" width="100%">
                    <tr style="height: 55px;">
                        <td>
                            <a href="ContentManagement/contentmanagement.aspx">
                                <img alt="help" src="images/icons/help_icon.gif" border="0" /></a>
                        </td>
                        <td>
                            <b><a title="Click " href="default.aspx"><span class="header">Website Info</span></a></b><span
                                class="headertext"><br />
                                Footer page content, Meta Data, Dynamic Content,Links etc.</span>
                        </td>
                    </tr>
                    <tr style="background-color: #eeeeee; height: 115px;">
                        <td align="left" colspan="2">
                            <ul class="listitem">
                                <!--li><a href="News/ManageNews.aspx">Manage News</a></li-->
                                <!--li><a href="MetaData/ManageMetaData.aspx">Manage MetaData </a></li-->
                                <li><a href="FooterContent\ManageFooterContent.aspx">Manage Footer Content</a></li>
                                <li><a href="DynamicPages\ManageDynamicPages.aspx">Manage Dynamic Content</a></li>
                                <br />
                                <br />
                                <li><a href="http://www.lamplight.com" target="_blank">Launch TikiBrand.com (B2C) Web
                                    Site</a></li>
                                <br />
                                <br />
                                <li><a href="http://sales.lamplight.com" target="_blank">Launch Sales.Lamplight.com
                                    (B2B) Web Site</a></li>
                                <li><a href="https://sales.lamplight.com/Manager" target="_blank">Launch sales.lamplight.com
                                    (B2B) Web Site Manager</a></li>
                                </span>
                            </ul>
                        </td>
                    </tr>
                </table>
            </td>
            <!-- End Manage Reports -->
            <!-- Begin Images-->
            <td valign="top" colspan="2" align="center">
                <table class="box" cellspacing="0" border="0" cellpadding="0" width="100%">
                    <tr style="height: 55px;">
                        <td>
                            <a href="default.aspx">
                                <img alt="" src="images/icons/imagetools.gif" border="0" /></a>
                        </td>
                        <td>
                            <b><a title="Click to Manage Your Products" href="Product/SelectProduct.aspx"><span
                                class="header">Manage Product </span></a></b>
                            <br />
                            <span class="headertext">Manage Product Description and Images</span>
                        </td>
                    </tr>
                    <tr style="background-color: #eeeeee; height: 115px;">
                        <td align="left" colspan="2">
                            <br />
                            <ul class="listitem">
                                <li><span class="listitem"><a href="Product/SelectProduct.aspx">Manage Product</a></span></li>
                                <li><span class="listitem"><a href="Product/DisabledWebItems.aspx">Manage Items with disabled "Add To Cart" button</a></span></li>
                                <li><span class="listitem"><a href="Product/unassignedimages.aspx">View Products without
                                    Images</a></span></li>
                                <li><span class="listitem"><a href="Product/managedefaultproductimage.aspx">Manage Web
                                    Site Default Product Image</a></span></li>
                                <li><span class="listitem"><a href="Product/BulkImageProcessing.aspx">Bulk Image Processing</a></span></li>
                                
                                <li><span class="listitem"><a href="Product/ProductRefresh.aspx">Order Motion Refresh</a></span></li>
                            </ul>
                            <asp:Label ID="lblLastProductImageManaged" runat="server"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td>
            <!-- End Images-->
            <!-- Begin Manage product -->
            <!--td valign="top" colspan="1" align="center" width="50%">
                <table class="box" cellspacing="0" cellpadding="0" rules="rows" width="100%">
                    <tr style="height:55px;">
                        <td>
                            <a href="default.aspx">
                                <img alt="" src="images/icons/retailer.gif" border="0" /></a>
                        </td>
                        <td>
                            <b><a title="Click to Manage Your Products" href="Product/ManageActiveMktgItems.aspx"><span class="header">
                                Manage Marketing Products</span></a></b><br />
                            <span class="headertext">Specify Products and supporting data for Marketing</span></td>
                    </tr>
                    <tr style="background-color:#eeeeee;height:115px;" >
                        <td align="left" colspan="2">
                            <br />
                            <ul>
                                <li><span class="listitem"><a href="Product/ManageActiveMktgItems.aspx">Manage Active Marketing Products</a></span></li>
                                <li><span class="listitem"><a href="Product/ManageCurrentYearProduct.aspx">Manage Current Year Products</a>
                                    (not currently used)</span></li>
                            </ul>
                        </td>
                    </tr>
                </table>
            </td-->
            <!-- End Manage product -->
        </tr>
        <tr>
            <td colspan="2">
                &nbsp;
            </td>
        </tr>
        <tr>
            <td>
                <table class="box" cellspacing="0" cellpadding="0" width="100%">
                    <tr style="height: 55px;">
                        <td>
                            <a href="ContentManagement/contentmanagement.aspx">
                                <img alt="" src="images/icons/icon_spec.jpg" border="0" /></a>
                        </td>
                        <td>
                            <b><a title="Click " href="ContentManagement/contentmanagement.aspx"><span class="header">
                                Content Management </span></a></b>
                            <br />
                            <span class="headertext">Management Home, Brand and Category DataCubes </span>
                        </td>
                    </tr>
                    <tr style="background-color: #eeeeee; height: 115px;">
                        <td align="left" colspan="2">
                            <ul class="listitem">
                                <li class="listitem"><a href="ContentManagement/managehomepage.aspx">Manage Datacubes</a></li>
                                <!--li class="listitem"><a href="ContentManagement/ManageSeriesDataCubes.aspx">Series Pages</a></li-->
                                <!--li class="listitem"><a href="ContentManagement/ManagePromotionDataCubes.aspx">Promotion Pages</a></li-->
                                <li class="listitem"><a href="Product/ManageProductSequence.aspx">Manage Product Sequence</a></li>
                            </ul>
                        </td>
                    </tr>
                </table>
            </td>
            <!-- begin promotions -->
            <td valign="top">
                <table class="box" cellspacing="0" cellpadding="0" width="100%">
                    <tr style="height: 55px;">
                        <td>
                            <a href="default.aspx">
                                <img alt="Product" src="images/icons/product.gif" border="0" /></a>
                        </td>
                        <td>
                            <b><a href="Promotion/ManagePromotions.aspx"><span class="header">Manage Media & Promotions</span></a></b><br />
                            <span class="headertext">Manage all of your News Information & Promotions.</span>
                        </td>
                    </tr>
                    <tr style="background-color: #eeeeee; height: 115px;">
                        <td align="left" colspan="2">
                            <ul class="listitem">
                                <li><span class="listitem"><a href="News/ManageNews.aspx">Manage News</a></span></li>
                                <li><span class="listitem"><a href="Promotion/ManagePromotions.aspx">Manage Promotions
                                </a></span></li>
                            </ul>
                        </td>
                    </tr>
                </table>
            </td>
            <!-- End promotions -->
            <!-- begin Recipes -->
            <!--td valign="top" colspan="1" align="center">
                <table class="box" cellspacing="0" cellpadding="0" rules="rows" width="100%">
                    <tr style="height:55px;">
                        <td>
                            <a href="default.aspx">
                                <img alt="" src="images/icons/product.gif" border="0" /></a>
                        </td>
                        <td>
                            <b><a title="Click to Manage Your Products" href="Recipes\ManageRecipes.aspx"><span class="header">
                                Manage Recipes</span></a></b><br />
                            <span class="headertext">Add, Update &amp; Delete Recipes</span></td>
                    </tr>
                    <tr style="background-color:#eeeeee;height:115px;">
                        <td align="left" colspan="2">
                            
                             <ul>
                                <li><span class="listitem"><a href="Recipes\ManageRecipes.aspx">Manage Recipes</a></span></li>
                            </ul>
                            
                            <asp:Label ID="lblLastRecipe" runat="server"></asp:Label>
                        </td>
                    </tr>
                </table>
            </td-->
            <!-- end Recipes -->
        </tr>
        <tr>
        <td valign="top" colspan="1" align="center" width="50%">
                <table class="box" cellspacing="0" cellpadding="0" rules="rows" width="100%">
                    <tr style="height: 55px;">
                        <td width="60px">
                            <a href="default.aspx">
                                <img alt="" src="images/icons/retailer.gif" border="0" /></a>
                        </td>
                        <td align="left">
                            <b><a title="Click to Manage Your Products" href="Products/Default.aspx"><span class="header">
                                Manage Site Redirects / Linkage / Tracking</span></a></b><br />
                            <span class="headertext">Manage internal search redirects, virtual folders and direct linkage/tracking</span>
                        </td>
                    </tr>
                    <tr style="background-color: #eeeeee; height: 115px;">
                        <td align="left" colspan="2" valign="top">
                            <br />
                             <asp:DataList ID="dlProductTools" Width="100%" runat="server"
                                DataSourceID="SiteMapDataSourceRedirects"  
                                RepeatColumns="2" HorizontalAlign="Left">
                                <HeaderTemplate>
                                    <ul >
                                </HeaderTemplate>
                                <ItemTemplate>
                                <li><span ><a href="<%# Eval("URL") %>"><%# Eval("Title") %> </a></span></li>
                                    
                                </ItemTemplate>
                                <ItemStyle VerticalAlign="Top" />
                                <FooterTemplate >
                                </ul>
                                </FooterTemplate>
                            </asp:DataList>
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
    </table>
 <asp:SiteMapDataSource ID="SiteMapDataSourceRedirects" runat="server" StartingNodeUrl="~/Redirects/default.aspx"
                SiteMapProvider="DefaultSiteMapProvider" ShowStartingNode="False"></asp:SiteMapDataSource>
</asp:Content>
