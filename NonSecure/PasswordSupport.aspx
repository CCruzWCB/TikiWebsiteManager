<%@ Page Language="VB" MasterPageFile="~/MasterPages/NonSecure.master" AutoEventWireup="false" CodeFile="PasswordSupport.aspx.vb" Inherits="Security_PasswordSupport" title="Untitled Page" %>
<%@ MasterType  VirtualPath ="~/MasterPages/NonSecure.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
    <asp:PasswordRecovery ID="PasswordRecovery1" runat="server" BackColor="#F7F7DE" BorderColor="#CCCC99"
        BorderStyle="Solid" BorderWidth="1px" Font-Names="Verdana" Font-Size="10pt">
        <TitleTextStyle BackColor="#6B696B" Font-Bold="True" ForeColor="White" />
    </asp:PasswordRecovery>

</asp:Content>

