
<%@ Page Language="VB" MasterPageFile="~/MasterPages/Home_Layout.master" AutoEventWireup="false"
    CodeFile="managecategorypage.aspx.vb" Inherits="ContentManagement_managecategorypage"
    Title="Untitled Page" %>
<%@ MasterType VirtualPath="~/MasterPages/Home_Layout.master" %>
<%@ Import Namespace="Management" %>

<%@ Register Src="../ClientWebControls/ClientDataCube.ascx" TagName="ClientDataCube"
    TagPrefix="uc2" %>
<asp:Content ID="Content1" runat="server" ContentPlaceHolderID="ContentPlaceHolder1">
    <uc2:ClientDataCube ID="ClientDataCube1" runat="server" Width="925" Height="123"
        ClassName="main" DataCubeID="" />
</asp:Content>
