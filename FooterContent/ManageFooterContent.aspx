<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="ManageFooterContent.aspx.vb" Inherits="FooterContent_ManageFooterContent" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ContentPlaceHolderID="contentMain" ID="ContentMain" runat="server">
   <div>
        <table border="0" cellspacing="0" cellpadding="0">
            <tr>
                <td width="800" rowspan="5" align="left" valign="top">
                    <table border="0" cellspacing="0" cellpadding="0" id="table1" onclick="return table1_onclick()">
                        <tr>
                            <td colspan="3">
                                <img src="../images/spacer.gif" width="12" height="12" alt="" /></td>
                        </tr>
                        <tr>
                            <td width="12">
                                <img src="../images/spacer.gif" width="12" height="12" alt="" /></td>
                            <td width="100%" align="left">
                                <!--body body body body body body body body body body body body body body body body body body body body body -->
                                <asp:ScriptManager ID="ScriptManager1" runat="server">
                                </asp:ScriptManager>
                                <table>
                                    <tr>
                                        <td class="headertd">
                                            Manage Static Footer Content</td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <cc1:CollapsiblePanelExtender ID="CollapsiblePanelExtender1" TargetControlID="CollapsablePanel1"
                                                CollapseControlID="imgExpandInfo" ExpandControlID="imgExpandInfo" runat="server">
                                            </cc1:CollapsiblePanelExtender>
                                            <asp:ImageButton ID="imgExpandInfo" ImageUrl="../images/icons/i-information.gif"
                                                runat="server" />
                                            <asp:Panel ID="CollapsablePanel1" runat="server">
                                                <span class="wiztxt">This page allows you to manage all static footer content.</span>
                                            </asp:Panel>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                     
                                            <asp:DropDownList ID="ddlDocument" runat="server" >
                                                <asp:ListItem Value="0">Select Document</asp:ListItem>
                                                <asp:ListItem Value="50">About Us </asp:ListItem>
                                                <asp:ListItem Value="46">Privacy Policy</asp:ListItem>
                                                <asp:ListItem Value="47">Employment Opportunities</asp:ListItem>
                                                <asp:ListItem Value="49">Intellectual Property Policy</asp:ListItem>
                                                <asp:ListItem Value="76">Return Policy</asp:ListItem>
                                                <asp:ListItem Value="33">Contact Us </asp:ListItem>
                                                <asp:ListItem Value="78">Mktg Email footer Text</asp:ListItem>
                                                <asp:ListItem Value="79">Help/Customer Service </asp:ListItem>
                                                 
                                            </asp:DropDownList><asp:Button ID="btnGo" runat="server" Text="Edit Document"></asp:Button></td>
                                    </tr>
                                   
                                </table>
                            </td>
                        </tr>
                        
                    </table>
                </td>
            </tr>
        </table>
    </div>

</asp:Content>