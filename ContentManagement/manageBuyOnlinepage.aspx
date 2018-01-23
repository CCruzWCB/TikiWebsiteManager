<%@ Register Src="../ClientWebControls/ClientDataCube.ascx" TagName="ClientDataCube"   TagPrefix="uc2" %>
<%@ Page Language="VB" MasterPageFile="~/MasterPages/BuyOnline_Layout.master" AutoEventWireup="false" CodeFile="manageBuyOnlinepage.aspx.vb" Inherits="MasterPages_manageBuyOnlinepage" title="Manage Buy Online Page content"  %>
<%@ MasterType VirtualPath ="~/MasterPages/BuyOnline_Layout.master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >
                 <uc2:ClientDataCube Width="451"  Height="360" ID="ClientDataCube1" EditMode="true"   DataCubeID="DataCubeBuy1" runat="server" />
</asp:Content>  
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder2" Runat="Server">
                 <uc2:ClientDataCube Width="240"  Height="120" ID="ClientDataCube2" EditMode="true" DataCubeID="DataCubeBuy2" runat="server" />
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="ContentPlaceHolder3" Runat="Server">
                 <uc2:ClientDataCube  Width="240"  Height="120" ID="ClientDataCube4" EditMode="true"  DataCubeID="DataCubeBuy3" runat="server" />
</asp:Content>    
<asp:Content ID="Content5" ContentPlaceHolderID="ContentPlaceHolder5" Runat="Server">
                 <uc2:ClientDataCube width="240" height="120" ID="ClientDataCube9" EditMode="true"  DataCubeID="DataCubeBuy4" runat="server" />
</asp:Content>  
