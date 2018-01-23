<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="ManageMetaData.aspx.vb" Inherits="MetaData_ManageMetaData" title="Manage MetaData" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">


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
                                            Manage News</td>
                                    </tr>
                                    <tr>
                                        <td>
                                            
                                            <asp:ImageButton ID="imgExpandInfo" ImageUrl="../images/icons/i-information.gif"
                                                runat="server" />
                                            <asp:Panel ID="CollapsablePanel1" runat="server">
                                                <span class="wiztxt">This page allows you to manage the meta data for the web site. </span>
                                            </asp:Panel>
                                        </td>
                                    </tr>
                                    
                                    <tr>
                                        <td>
                                            <asp:Label ID="lblMessage" runat="server" Visible="True"></asp:Label></td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                        
                        <tr>
                            <td colspan="3"><br /><br /><br /><asp:FormView ID="fvKeywords" runat="server" DataSourceID="dsMetaDataKeywords" DefaultMode="Edit">
                                <EditItemTemplate>
                                    <b>Meta Keywords</b><br />
                                    <asp:TextBox Width="600" TextMode="MultiLine" Rows="6"  ID="txtKeywords" runat="server" Text='<%# eval("meta_content") %>'>
                                    </asp:TextBox><br />
                                    <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" CommandName="Update"
                                        Text="Update">
                                    </asp:LinkButton>
                                    <asp:LinkButton ID="UpdateCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                                        Text="Cancel">
                                    </asp:LinkButton>
                                </EditItemTemplate>
                              
                                </asp:FormView>
                            </td>
                        </tr>
                          <tr>
                            <td colspan="3"><br /><br /><br /><asp:FormView ID="fvDescription" runat="server" DataSourceID="dsMetaDataDescription" DefaultMode="Edit">
                                <EditItemTemplate>
                                    <b>Meta Description</b><br />
                                    <asp:TextBox Width="600" TextMode="MultiLine" Rows="6"  ID="txtDescription" runat="server" Text='<%# eval("meta_content") %>'>
                                    </asp:TextBox><br />
                                    <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" CommandName="Update"
                                        Text="Update">
                                    </asp:LinkButton>
                                    <asp:LinkButton ID="UpdateCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                                        Text="Cancel">
                                    </asp:LinkButton>
                                </EditItemTemplate>
                              
                                </asp:FormView>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="3">
                                &nbsp;
                         
                                &nbsp;
                                <asp:SqlDataSource ID="dsMetaDataKeywords" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
                                    SelectCommand="[GetMetaData]" SelectCommandType="StoredProcedure" UpdateCommand="UpdateMetaDataKeywords" UpdateCommandType="storedProcedure" >
                                    <SelectParameters>
                                        <asp:Parameter DefaultValue="keywords" Name="meta_name" Type="string" />
                                    </SelectParameters>
                                    <UpdateParameters>
                                        <asp:controlParameter Name="meta_keywords" Type="string" Size="1000" ControlID="fvKeywords$txtKeywords" PropertyName="Text"  />
                                    </UpdateParameters>
                                </asp:SqlDataSource>
                                
                                  <asp:SqlDataSource ID="dsMetaDataDescription" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
                                    SelectCommand="[GetMetaData]" SelectCommandType="StoredProcedure" UpdateCommand="UpdateMetaDataDescription" UpdateCommandType="storedProcedure" >
                                    <SelectParameters>
                                        <asp:Parameter DefaultValue="description" Name="meta_name" Type="string" />
                                    </SelectParameters>
                                    <UpdateParameters>
                                        <asp:controlParameter Name="meta_description" Type="string" Size="1000" ControlID="fvDescription$txtDescription" PropertyName="Text"  />
                                    </UpdateParameters>
                                </asp:SqlDataSource>
                               
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </div>


</asp:Content>

