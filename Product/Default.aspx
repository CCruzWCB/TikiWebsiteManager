<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="MasterPages_Default" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">

Enter Product Number: <asp:TextBox ID="txtProductNumber" runat="server"></asp:TextBox><asp:Button ID="btnSearch" runat="server" Text="Search" />

Or 
<br />
Browse: 
    <asp:TreeView runat="server" ID="ProductTree" ></asp:TreeView>




</asp:Content>

