<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="selectproduct.aspx.vb" Inherits="Product_selectproduct" title="Untitled Page" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server" >

<asp:Panel ID="pnlMain" runat="server" DefaultButton="btnGo">
    Enter Product Number
    <br />
    <asp:TextBox ID="txtProductNumber" runat="server"></asp:TextBox>&nbsp;<asp:Button
        ID="btnGo" runat="server" Text="Go" /><br />
    <br />
    Or
    <br />
    <br />
    <asp:DropDownList ID="ddlProducts" runat="server" AppendDataBoundItems="True" AutoPostBack="True"
        DataSourceID="dsProducts" DataTextField="Description" DataValueField="ProductID">
        <asp:ListItem Selected="True" Value="0">Select Product</asp:ListItem>
    </asp:DropDownList><asp:SqlDataSource ID="dsProducts" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
        SelectCommand="GetAllProducts" SelectCommandType="StoredProcedure">
        <SelectParameters >
        <asp:Parameter Name="BrandID"  Type ="Int32" DefaultValue ='<%$ AppSettings:BrandID %>' />
        </SelectParameters>
        </asp:SqlDataSource>
    &nbsp;

</asp:Panel>

</asp:Content>

