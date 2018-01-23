<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="ManageDynamicPages.aspx.vb" Inherits="FooterContent_ManageFooterContent" %>
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
                                            Manage Dynamic Page Content</td>
                                    </tr>
                                    <tr>
                                        <td>
                                            <cc1:CollapsiblePanelExtender ID="CollapsiblePanelExtender1" TargetControlID="CollapsablePanel1"
                                                CollapseControlID="imgExpandInfo" ExpandControlID="imgExpandInfo" runat="server">
                                            </cc1:CollapsiblePanelExtender>
                                            <asp:ImageButton ID="imgExpandInfo" ImageUrl="../images/icons/i-information.gif"
                                                runat="server" />
                                            <asp:Panel ID="CollapsablePanel1" runat="server">
                                                <span class="wiztxt">This page allows you to manage all dynamic page content.</span>
                                            </asp:Panel>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                        <br /><br />
                                        Add New Dynamic Page <br />
                                        Name:  <asp:TextBox ID="txtName" MaxLength="100" runat="server"></asp:TextBox>  <asp:Button ID="btnAdd" runat="server" Text="Add" />
                                        </td>
                                    </tr>
                                    
                                    <tr>
                                        <td>
                                     
                                            
                                            <asp:GridView ID="gvDocs" runat="server" AutoGenerateColumns="False" DataKeyNames="doc_id"
                                                DataSourceID="dsDocList">
                                                <Columns>
                                                    <asp:BoundField DataField="doc_id" HeaderText="doc_id" InsertVisible="False" ReadOnly="True"
                                                        SortExpression="doc_id" Visible="False" />
                                                    
                                                    <asp:TemplateField HeaderText="doc_name" SortExpression="doc_name">
                                                        <EditItemTemplate>
                                                            <asp:TextBox ID="txtDocName" MaxLength="100" runat="server" Text='<%# Bind("doc_name") %>'></asp:TextBox>
                                                        </EditItemTemplate>
                                                        <ItemTemplate>
                                                            <asp:HyperLink ID="hlEditDoc" runat="server" Text='<%# Bind("doc_name") %>' NavigateUrl='<%# Eval("doc_id", "EditDocument.aspx?doc_id={0}") %>'></asp:HyperLink>
                                                        </ItemTemplate>
                                                        
                                                    </asp:TemplateField>
                                                    <asp:TemplateField HeaderText="Doc URL"  >
                                                        <ItemTemplate>
                                                            <asp:HyperLink ID="hlURL" ToolTip ="Click to view document in new window" Target ="_blank"   runat="server" Text='<%# String.Format("{0}/Content/{1}/{2}.aspx",ConfigurationManager.AppSettings("ApplicationHTTPRoot"),Eval("doc_id"), Eval("doc_name")) %>' NavigateUrl='<%# String.Format("{0}/Content/{1}/{2}.aspx",ConfigurationManager.AppSettings("ApplicationHTTPRoot"),Eval("doc_id"), Eval("doc_name")) %>'></asp:HyperLink>
                                                        </ItemTemplate>
                                                        
                                                    </asp:TemplateField>
                                                    <asp:BoundField DataField="description"   HeaderText="description" SortExpression="description" />
                                                    <asp:CommandField ShowEditButton="True" />
                                                    <asp:CommandField ShowDeleteButton="True" />
                                                </Columns>
                                            </asp:GridView>
                                            <br />
                                            <asp:SqlDataSource ID="dsDocList" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
                                                SelectCommand="GetDocumentsByCatandSub" SelectCommandType="StoredProcedure" UpdateCommand="[UpdateDocumentName]" UpdateCommandType="StoredProcedure" InsertCommand="AddDynamicDocument" InsertCommandType="StoredProcedure" 
                                                 DeleteCommand="DeleteDocument" DeleteCommandType="StoredProcedure"  >
                                                <DeleteParameters>
                                                <asp:ControlParameter Name="doc_id" ControlID="gvDocs" PropertyName="SelectedValue" />
                                                </DeleteParameters>
                                                <UpdateParameters>
                                                    <asp:ControlParameter Name="doc_id" ControlID="gvDocs" PropertyName="SelectedValue" />
                                                    <asp:Parameter Name="doc_name" Type="String" />
                                                    <asp:Parameter Name="description" Type="String" />
                                                </UpdateParameters>
                                                <SelectParameters>
                                                    <asp:Parameter DefaultValue="9" Name="doc_cat_id" Type="Int32" />
                                                    <asp:Parameter DefaultValue="1" Name="doc_sub_id" Type="Int32" />
                                                </SelectParameters>
                                                <InsertParameters>
                                                    <asp:Parameter Name="doc_name" Type="String" />
                                                    <asp:Parameter Direction="InputOutput" Name="doc_id" Type="Int32" />
                                                </InsertParameters>
                                            </asp:SqlDataSource>
                                            </td>
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