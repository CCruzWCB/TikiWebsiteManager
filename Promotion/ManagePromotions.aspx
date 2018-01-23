<%@ Page Language="VB" AutoEventWireup="false"  MasterPageFile="~/MasterPages/MasterPage.master" CodeFile="ManagePromotions.aspx.vb" Inherits="Promotion_ManagePromotions" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ContentPlaceHolderID="contentMain" ID="ContentMain" runat="server" >

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
                                            Manage Promotions</td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <cc1:CollapsiblePanelExtender ID="CollapsiblePanelExtender1" TargetControlID="CollapsablePanel1"
                                                CollapseControlID="imgExpandInfo" ExpandControlID="imgExpandInfo" runat="server">
                                            </cc1:CollapsiblePanelExtender>
                                            <asp:ImageButton ID="imgExpandInfo" ImageUrl="../images/icons/i-information.gif"
                                                runat="server" />
                                            <asp:Panel ID="CollapsablePanel1" runat="server">
                                                <span class="wiztxt">This page allows you to manage all Promotions on the web site. To
                                                    add a new Promotion, click the "Add new Promotion" link below. To edit an existing Promotion,
                                                    simply click on the Promotion name in the list below. </span>
                                            </asp:Panel>
                                        </td>
                                    </tr>
                                    
                                    <tr>
                                        <td class="wizhdr">
                                            <br />
                                            To add a Promotion,
                                            <asp:HyperLink ID="HyperLinkAdd" title="Click To Add Promotion" runat="server" NavigateUrl="addEditPromotion.aspx">Click Here</asp:HyperLink></td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <asp:Label ID="lblMessage" runat="server" Visible="True"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3">
                                &nbsp;
                                <asp:GridView ID="gvPromotions" runat="server" AutoGenerateColumns="False" 
                                    GridLines="Vertical" PageSize="20" AllowPaging="True" AllowSorting="True" DataKeyNames="promotionID" DataSourceID="dsPromotions">
                                    <Columns>
                                        <asp:TemplateField HeaderText="PromotionID" Visible="False">
                                            <ItemTemplate>
                                                <asp:Label ID="lblPromotionID" runat="server" Text='<%# Bind("PromotionID") %>'>
                                                </asp:Label>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Promotion Name">
                                            <ItemTemplate>
                                                <asp:HyperLink ID="HyperLink1" runat="server" Target="_self" Text='<%# Bind("Name") %>'
                                                    NavigateUrl='<%# "AddEditPromotion.aspx?pid=" & Eval("PromotionID") %>'></asp:HyperLink>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Promotion URL"  >
                                                        <ItemTemplate>
                                                            <asp:HyperLink ID="hlURL" ToolTip ="Click to view promotion in new window" Target ="_blank"   runat="server" Text='<%# String.Format("{0}/Consumer/viewpromotion.aspx?PromotionID={1}",ConfigurationManager.AppSettings("ApplicationHTTPRoot"),Eval("PromotionID")) %>' NavigateUrl='<%# String.Format("{0}/Consumer/viewpromotion.aspx?PromotionID={1}",ConfigurationManager.AppSettings("ApplicationHTTPRoot"),Eval("PromotionID")) %>'></asp:HyperLink>
                                                        </ItemTemplate>
                                                        
                                                    </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Status">
                                            <ItemTemplate>
                                                <asp:Label ID="lblStatus" runat="server" Text='<%# ConvertStatus(eval("Active")) %>'></asp:Label>
                                            </ItemTemplate>
                                            <EditItemTemplate>
                                                <asp:TextBox ID="TextBox1" runat="server" Text='<%# Bind("Active") %>'>
                                                </asp:TextBox>
                                            </EditItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField HeaderText="Data Cube">
                                            <ItemTemplate>
                                                <asp:HyperLink ID="hlDataCube" runat="server" Target="_self" Text='Manage Data Cube'
                                                    NavigateUrl='<%# "~/ContentManagement/ManagePromotionDataCubes.aspx?PromotionID=" & Eval("PromotionID") %>'></asp:HyperLink>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                        <asp:TemplateField>
                                            <ItemTemplate>
                                                <asp:LinkButton ID="btnDelete" CommandArgument='<%# Eval("PromotionID") %>' OnClientClick="return confirm('You are about to delete this Promotion. Are you sure?');"
                                                    runat="server" CausesValidation="false" CommandName="Delete" Text="Delete"></asp:LinkButton>
                                            </ItemTemplate>
                                        </asp:TemplateField>
                                    </Columns>
                                </asp:GridView>
                                <asp:SqlDataSource ID="dsPromotions" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
                                    DeleteCommand="DeletePromotion" DeleteCommandType="StoredProcedure" SelectCommand="GetPromotion"
                                    SelectCommandType="StoredProcedure">
                                    <DeleteParameters>
                                        <asp:Parameter Name="PromotionID" Type="Int32" />
                                    </DeleteParameters>
                                    <SelectParameters>
                                        <asp:Parameter DefaultValue="0" Name="PromotionID" Type="Int32" />
                                    </SelectParameters>
                                </asp:SqlDataSource>
                                &nbsp;
                                
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>


</asp:Content>