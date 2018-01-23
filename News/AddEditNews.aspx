<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="AddEditNews.aspx.vb" Inherits="News_AddEditNews" title="Untitled Page" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">

  <asp:Panel ID="pNewsInfo" runat="server" Visible="True">
        <table width="600" border="0">
            <tr>
                <td colspan="2">
                    <span class="wiztxt">This page allows you add or edit an existing news article. Complete the
                        information for the article.</span>
                </td>
            </tr>
            <tr bgcolor="#eeeee0">
                <td align="left">
                    News Information</td>
                
                <td>
                    <asp:Label ID="lblMessage" runat="server" ForeColor="red"></asp:Label></td>
            </tr>
        </table>
        <asp:FormView ID="fvNews" runat="server" DataSourceID="dsNews" 
            DefaultMode="Edit"   >
            <InsertItemTemplate>
            <table border="0">
                    <tr>
                        <td>
                            Title:</td>
                        <td colspan="1" align="left">
                            <asp:TextBox ID="txtTitle" runat="server" Text='<%# Bind("Title") %>'>
                            </asp:TextBox>
                            <asp:RequiredFieldValidator ID="rfvTitle" runat="server" ErrorMessage="Title Required"
                                ControlToValidate="txtTitle">
                            </asp:RequiredFieldValidator></td>
                        <td align="center" rowspan="3" >&nbsp;
                            <</td>
                    </tr>
                    <tr>
                        <td>
                            Description:</td>
                        <td colspan="2" align="left">
                            <asp:TextBox ID="txtDescription" runat="server" Text='<%# Bind("description") %>'>
                            </asp:TextBox>
                            <asp:RequiredFieldValidator ID="rfvDescription" runat="server" ErrorMessage="Description Required"
                                ControlToValidate="txtDescription">
                            </asp:RequiredFieldValidator></td>
                    </tr>
                   
                    <tr>
                        <td align="center">
                            Enabled:
                            <asp:CheckBox ID="CheckboxEnabled" runat="server" Checked='<%# Bind("active") %>'></asp:CheckBox></td>
                        <td align="center">
                            </td>
                    </tr>
                   
                    <tr>
                        <td>
                            File:
                        </td>
                        <td colspan="2" align="left">
                        <asp:FileUpload ID="FileUpload" runat="server"  />
                          </td>
                    </tr>
                 
                    <tr>
                        <td align="center" colspan="3">
                            <asp:Button ID="btnCancelAdd" runat="server" Text="Cancel" PostBackUrl="ManageNews.aspx" CausesValidation="False">
                            </asp:Button>
                            <asp:Button ID="btnAddRecipe" runat="server" CommandName="Insert" Text="Save"></asp:Button>
                            </td>
                    </tr>
                </table>
            </InsertItemTemplate>
            <EditItemTemplate>
            <table border="0">
                    <tr>
                        <td>
                            Title:</td>
                        <td colspan="1" align="left">
                            <asp:TextBox ID="txtTitle" runat="server" Text='<%# Bind("Title") %>'>
                            </asp:TextBox>
                            <asp:RequiredFieldValidator ID="rfvTitle" runat="server" ErrorMessage="Title Required"
                                ControlToValidate="txtTitle">
                            </asp:RequiredFieldValidator></td>
                        <td align="center" rowspan="3" ></td>
                    </tr>
                  
                    <tr>
                        <td>
                            Description:</td>
                        <td colspan="2" align="left">
                            <asp:TextBox ID="txtDescription" runat="server" Text='<%# Bind("description") %>'>
                            </asp:TextBox>
                            <asp:RequiredFieldValidator ID="rfvDescription" runat="server" ErrorMessage="Description Required"
                                ControlToValidate="txtDescription">
                            </asp:RequiredFieldValidator></td>
                    </tr>
                  
                    <tr>
                        <td align="center">
                            Enabled:
                            <asp:CheckBox ID="CheckboxEnabled" runat="server" Checked='<%# Bind("active") %>'></asp:CheckBox></td>
                        <td align="center">
                            </td>
                    </tr>
                   
                    <tr>
                        <td>
                            File:
                        </td>
                        <td colspan="2" align="left">
                        <asp:FileUpload ID="FileUpload" runat="server" /><br />
                        Current File:  <asp:Label ID="lblFile" runat="server" Text='<%# Bind("LocalSystemPath") %>'></asp:Label>
                        
                          </td>
                    </tr>
                  
                    <tr>
                        <td align="center" colspan="3">
                            <asp:Button ID="btnCancelAdd" runat="server" Text="Cancel" CausesValidation="False" PostBackUrl="ManageNews.aspx">
                            </asp:Button>
                            <asp:Button ID="btnUpdateRecipe" runat="server" CommandName="Update" Text="Update"></asp:Button>
                            </td>
                    </tr>
                </table>
            
            </EditItemTemplate>
        </asp:FormView>
    </asp:Panel>
    <table>
        
    </table>
    <asp:SqlDataSource ID="dsNews" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
        SelectCommand="GetKnowledge" SelectCommandType="StoredProcedure" UpdateCommand="[EditNewsRoomArticle]" UpdateCommandType="StoredProcedure"
         InsertCommand="[AddNewsRoomArticle]" InsertCommandType="StoredProcedure" >
     
        <SelectParameters>
            <asp:QueryStringParameter DefaultValue="0" Name="KnowledgeID" Type="Int32" QueryStringField="kid" />
        </SelectParameters>
        <DeleteParameters>
            <asp:QueryStringParameter DefaultValue="0" Name="recipe_id" Type="Int32" QueryStringField="rid" />
        </DeleteParameters>
        <InsertParameters>
            <asp:Parameter Name="Title" Type="String" />
            <asp:Parameter Name="Description" Type="String" />
            <asp:Parameter Name="LocalSystemPath" Type="String" />
            <asp:Parameter Name="Active" Type="Boolean" />
            <asp:Parameter Name="KnowledgeID" Type="int32" Direction="Output" /> 
        </InsertParameters>
        <UpdateParameters>
            <asp:QueryStringParameter Name="KnowledgeID" QueryStringField="kid" DefaultValue="0" Type="Int32"  />
            <asp:Parameter Name="Title" Type="String" />
            <asp:Parameter Name="Description" Type="String" />
            <asp:Parameter Name="LocalSystemPath" Type="String" />
            <asp:Parameter Name="Active" Type="Boolean" />
     
        </UpdateParameters>
    </asp:SqlDataSource>
   

</asp:Content>

