<%@ Page Language="VB" MasterPageFile="~/MasterPages/Home_Layout1.master" AutoEventWireup="false" CodeFile="ManageBrandDataCubes.aspx.vb" Inherits="ContentManagement_ManageBrandDataCubes" title="Untitled Page" %>
<%@ MasterType VirtualPath="~/MasterPages/Home_Layout1.master" %>
<%@ Import Namespace="Management" %>
<%@ Register Src="../ClientWebControls/ClientDataCube.ascx" TagName="ClientDataCube"
    TagPrefix="uc1" %>
  
<asp:Content ID="contentMain" runat="server" ContentPlaceHolderID="ContentPlaceHolder1">


  <table cellpadding ="0" cellspacing ="0" width ="900" border="0"><tr>
<td valign ="top">
<uc1:ClientDataCube ID="DataCube1" EditMode="True" runat="server" Height="398" Width="664" />
<uc1:ClientDataCube ID="DataCube5" EditMode="True" runat="server" Height="82" Width="664" />
</td>
<td valign ="top"  width="300">
    <table cellpadding ="0" cellspacing="0" border ="0">
    <tr><td><uc1:ClientDataCube ID="DataCube2" EditMode="True" runat="server" Height="163" Width="286" /></td></tr>
    <tr><td><uc1:ClientDataCube ID="DataCube3" EditMode="True" runat="server" Height="159" Width="286" /></td></tr>
    <tr><td><uc1:ClientDataCube ID="DataCube4" EditMode="True" runat="server" Height="158" Width="286" /></td></tr>
    </table>
</td></tr>
</table>
    <table cellpadding ="0" cellspacing="0" width="900">
<tr><td width="300"></td>
<td width="300"></td></tr>
</table>

</asp:Content>
