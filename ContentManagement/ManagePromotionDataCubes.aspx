<%@ Page Language="VB" MasterPageFile="~/MasterPages/Promo_Layout.master" AutoEventWireup="false" CodeFile="ManagePromotionDataCubes.aspx.vb" Inherits="ContentManagement_ManagePromotionDataCubes" title="Untitled Page" %>
<%@ MasterType VirtualPath="~/MasterPages/Promo_Layout.master" %>
<%@ Import Namespace="Management" %>
<%@ Register Src="../ClientWebControls/ClientDataCube.ascx" TagName="ClientDataCube"
    TagPrefix="uc2" %>
    
  
<asp:Content ID="contentMain" runat="server" ContentPlaceHolderID="ContentPlaceHolder1">


    <uc2:ClientDataCube ID="ClientDataCube1" runat="server" Width="687" Height="422"
        ClassName="main" DataCubeID="" />
</asp:Content>
