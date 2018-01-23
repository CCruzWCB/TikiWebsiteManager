<%@ Page Theme ="Management"  Language="VB" MasterPageFile="~/MasterPages/NonSecure.master" AutoEventWireup="false" CodeFile="Login.aspx.vb" Inherits="Security_Login" title="Login" %>
<%@ MasterType VirtualPath ="~/MasterPages/NonSecure.master" %>
 
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
<asp:Panel ID="pLogin" runat="server" DefaultButton ="btnLogin">
 <table cellspacing="0" cellpadding="1" border="0" id="ctl00_ContentMain_Login1" style="background-color:#F7F7DE;border-color:#CCCC99;border-width:1px;border-style:Solid;border-collapse:collapse;">
	<tr>
		<td><table cellpadding="0" border="0" style="font-family:Verdana;font-size:10pt;">
			<tr>
				<td align="center" colspan="2" style="color:White;background-color:#6B696B;font-weight:bold;">Log In</td>
			</tr><tr>
				<td align="right"><label for="ctl00_ContentMain_Login1_UserName">User Name:</label></td><td><asp:TextBox ID="txtUserName" runat="server"  ></asp:TextBox></td>
			</tr><tr>
				<td align="right"><label for="ctl00_ContentMain_Login1_Password">Password:</label></td><td><asp:TextBox ID="txtPassword"  TextMode ="Password"  runat ="server" autocomplete="off" ></asp:TextBox></td>
			</tr><tr>
				<td colspan="2"><asp:CheckBox ID="chkRememberMe" runat="server" Text ="Remember me on the next login" /></td>
			</tr><tr>
				<td align="right" colspan="2" style="height: 24px"><asp:Button ID="btnLogin" runat="server" Text ="Log In" /></td>
			</tr>
		</table></td>
	</tr>
</table>
    </asp:Panel>
</asp:Content>
