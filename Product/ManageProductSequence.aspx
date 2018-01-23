<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="ManageProductSequence.aspx.vb" Inherits="ContentManagement_ManageProductSequence" title="Manage Product Sequence" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>


<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server" >

    <asp:ScriptManager ID="ScriptManager" runat="server"></asp:ScriptManager>
    <div style="width: 70%; margin: 5px;" class="box">
                    <div style="margin: 5px;" class="graybox">
                        Sequence Selection:
                        <asp:DropDownList ID="ddProductSeries" runat="server" DataSourceID="dsGetProductCategorySequence"
                            DataValueField="SortOrder" DataTextField="Categorization" AppendDataBoundItems="true"
                            AutoPostBack="true">
                            <asp:ListItem Value="0" Text="Select one"></asp:ListItem>
                        </asp:DropDownList>
                    </div>
                    <br />
                    <br />







                    <!-- Begin Sub Category Listing for Sorting -->
                    <asp:Panel ID="pSubCategories" runat="server" Visible="false">
                        <div style="width: 400px;" class="reorderListContainer">
                            <b>Manage Category Sequence </b>
                            <cc1:ReorderList ID="rlSubCategoryItems" runat="server" ItemInsertLocation="Beginning"
                                DragHandleAlignment="Left" AllowReorder="true" DataKeyField="ProductCategoryID"
                                PostBackOnReorder="False" DataSourceID="dsGetSubCategoryListingsbySequence" SortOrderField="Sequence"
                                CallbackCssStyle="callbackStyle">
                                <ReorderTemplate>
                                 <div class="reorderCue" style="height:20px;width:400px;">
                                    </div>
                                    <%--<asp:Panel ID="Panel2" runat="server" Width="400" Height="20" CssClass="reorderCue" />--%>
                                </ReorderTemplate>
                                <EmptyListTemplate>
                                    There are currently no Items for sorting
                                </EmptyListTemplate>
                                <DragHandleTemplate>
                                    <div class="dragHandle">
                                    </div>
                                </DragHandleTemplate>
                                <ItemTemplate>
                                    <table id="tblPromo" runat="server" width="400" style="text-align: left; border: solid 1px black;">
                                        <tr align="left" style="">
                                            <td width="80">
                                                <asp:Image ID="imgSubCategory" Width="75" Height="75" runat="server" ImageUrl='<%# Utilities.GetImageURL(Eval("ImagePath"))%>' />
                                            </td>
                                            <td>
                                                <%#Eval("Name")%>
                                            </td>
                                        </tr>
                                    </table>
                                </ItemTemplate>
                            </cc1:ReorderList>
                        </div>
                    </asp:Panel>
                    <!-- End  Sub Category Listing for Sorting -->
                    <!-- Begin Product Listing for Sorting -->
                    <asp:Panel ID="pProducts" runat="server" Visible="false">
                        <div style="width: 400px;" class="reorderListContainer">
                            <b>Manage Product Sequence </b>
                            <cc1:ReorderList ID="rlProducts" runat="server" ItemInsertLocation="Beginning" DragHandleAlignment="Left"
                                AllowReorder="true" DataKeyField="ProductID" PostBackOnReorder="False" DataSourceID="dsGetProductListingBySequence"
                                SortOrderField="Sequence" CallbackCssStyle="callbackStyle">
                                <ReorderTemplate>
                                    <asp:Panel ID="Panel2" runat="server" Width="400" Height="20" CssClass="reorderCue" />
                                </ReorderTemplate>
                                <EmptyListTemplate>
                                    There are currently no Items for sorting
                                </EmptyListTemplate>
                                <DragHandleTemplate>
                                    <div class="dragHandle">
                                    </div>
                                </DragHandleTemplate>
                                <ItemTemplate>
                                    <table id="tblPromo" runat="server" width="400" style="text-align: left; border: solid 1px black;">
                                        <tr align="left" style="">
                                            <td width="80">
                                                <asp:Image ID="imgProduct" Width="75" Height="75" runat="server" ImageUrl='<%# Utilities.GetImageURL(Eval("ImagePath"))%>' />
                                            </td>
                                            <td>
                                                <%#Eval("Name")%>
                                            </td>
                                        </tr>
                                    </table>
                                </ItemTemplate>
                            </cc1:ReorderList>
                        </div>
                    </asp:Panel>
                    <!-- End  Product Listing for Sorting -->
                </div>
    <asp:SqlDataSource ID="dsGetProductCategorySequence" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
        SelectCommand="[Manager].[GetAllCategorizationSequence]" SelectCommandType="StoredProcedure">

        <SelectParameters>
            <asp:Parameter DefaultValue='<%$ AppSettings:BrandID %>' Type="Int32" Name="BrandID" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsGetSubCategoryListingsbySequence" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
        SelectCommand="[Manager].[GetSubCategoryListingsBySequence]" SelectCommandType="StoredProcedure"
        UpdateCommand="[Manager].UpdateSubCategorySequence" UpdateCommandType="StoredProcedure">
        <UpdateParameters>
            <asp:Parameter Name="ProductCategoryID" Type="int32" />
            <asp:Parameter Name="Sequence" Type="Int32" />
        </UpdateParameters>
        <SelectParameters>
            <asp:Parameter Name="ParentCategoryID" Type="Int32" />
        </SelectParameters>
    </asp:SqlDataSource>
    <asp:SqlDataSource ID="dsGetProductListingBySequence" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
        SelectCommand="[Manager].[GetProductListingsBySequence]" SelectCommandType="StoredProcedure"
        UpdateCommand="[Manager].UpdateProductSequence" UpdateCommandType="StoredProcedure">
        <UpdateParameters>
            <asp:Parameter Name="ProductID" Type="Int32" />
            <asp:Parameter Name="Sequence" Type="Int32" />
            
        </UpdateParameters>
        <SelectParameters>
            <asp:Parameter Name="ProductCategoryID" Type="Int32" />
            <asp:Parameter DefaultValue='<%$ AppSettings:BrandID %>' Type="Int32" Name="BrandID" />
        </SelectParameters>
    </asp:SqlDataSource>
    
</asp:Content>
<%--

<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ManageProductSequence.aspx.vb"
    Inherits="ContentManagement_ManageProductSequence" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Untitled Page</title>
</head>
<body style="margin-left: 0; margin-top: 0; margin-right: 0">
    <form id="form1" runat="server">
    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    <table border="0" cellpadding="3" cellspacing="0" width="100%">
        <tr>
            <td class="pageheader">
                Manage Product Sequence
            </td>
        </tr>
    </table>
    <table width="100%" border="0" cellpadding="5" cellspacing="0">
        <tr>
            <td style="width: 100%" align="center" valign="middle">

            </td>
        </tr>
    </table>
    </form>
</body>
</html>
--%>