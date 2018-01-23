<%@ Register Src="../ClientWebControls/ClientDataCube.ascx" TagName="ClientDataCube"   TagPrefix="uc1" %>
<%@ Page Language="VB" MasterPageFile="~/MasterPages/Home_Layout.master" AutoEventWireup="false" CodeFile="managehomepage.aspx.vb" Inherits="MasterPages_managehomepage" title="Manage Home Page content"  %>
<%@ MasterType VirtualPath ="~/MasterPages/Home_Layout.master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server" >

<table cellpadding ="0" cellspacing ="0" width ="900" border="0"><tr>
<td valign ="top">
<uc1:ClientDataCube ID="DataCube0_1" EditMode="true"    DataCubeID="DataCube0_1" runat="server" Height="398" Width="664" />

<uc1:ClientDataCube ID="DataCube0_5" EditMode="true"    DataCubeID="DataCube0_5" runat="server" Height="82" Width="664" />
</td>
<td valign ="top"  width="300">
    <table cellpadding ="0" cellspacing="0" border ="0">
    <tr><td><uc1:ClientDataCube ID="DataCube0_2" EditMode="true" DataCubeID="DataCube0_2" runat="server" Height="163" Width="286" /></td></tr>
    <tr><td><uc1:ClientDataCube ID="DataCube0_3" EditMode="true" DataCubeID="DataCube0_3" runat="server" Height="159" Width="286" /></td></tr>
    <tr><td><uc1:ClientDataCube ID="DataCube0_4" EditMode="true" DataCubeID="DataCube0_4" runat="server" Height="158" Width="286" /></td></tr>
    </table>
</td></tr>
</table>
    <table cellpadding ="0" cellspacing="0" width="900">
<tr><td width="300"></td>
<td width="300"></td></tr>
</table>

</asp:Content>  
