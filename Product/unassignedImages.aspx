<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="unassignedImages.aspx.vb" Inherits="Product_unassignedImages" title="Untitled Page" %>
<%@ MasterType VirtualPath ="~/MasterPages/MasterPage.master" %>
<%@ Import Namespace ="Management" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">

<asp:DataList ID="dlMissingProductImages" runat ="server" DataSourceID="dsMissingProductsImages" RepeatColumns="4" RepeatDirection="Horizontal" >
    <ItemTemplate>
        <asp:ImageButton  runat ="server" CommandName ="Select" CommandArgument ='<%# Eval("ProductID") %>' AlternateText ="Click to Assign Image" ImageUrl ='<%# Utilities.GetImageURL(Eval("ImagePath_small")) %>' />
        <br />Category:
        <asp:Label ID="ProductCategoryNameLabel" runat="server" Text='<%# Eval("ProductCategoryName") %>'>
        </asp:Label><br />
        ProductID:
        <asp:Label ID="ProductSeriesIDLabel" runat="server" Text='<%# Eval("ProductID") %>'>
        </asp:Label><br />
        <a  target ="_blank" href ='/Consumer/product_detail_e.aspx?ProductID=<%# Eval("ProductID") %>'>Launch Product Page</a><br />
        <asp:Label ID="NameLabel" runat="server" Text='<%# Eval("Name") %>'></asp:Label><br />
        Model Number:
        <asp:Label ID="ProductSeriesNumberLabel" runat="server" Text='<%# Eval("ProductModelNumber") %>'>
        </asp:Label><br />
        <br />
    </ItemTemplate>
</asp:DataList><asp:SqlDataSource ID="dsMissingProductsImages" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
    SelectCommand="GetProductswithoutPrimaryImages" SelectCommandType="StoredProcedure">
    <SelectParameters >
    <%--<asp:Parameter Name="BrandID"  Type ="Int32" DefaultValue ='<%$ AppSettings:BrandID %>' />--%>
    </SelectParameters>
</asp:SqlDataSource>

</asp:Content>

