<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="EditDocument.aspx.vb" Inherits="FooterContent_EditDocument" title="Untitled Page" validateRequest="false" %>
<%@ Register TagPrefix="cc1" Namespace="Management" %>
   <%@ Register TagPrefix="FTB" Namespace="FreeTextBoxControls" Assembly="FreeTextBox" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ContentPlaceHolderID="contentMain" ID="ContentMain" runat="server" >
    <asp:Panel ID="pnlDocumentInfo" runat="server" Visible="True">
        <table width="800" border="0">
            <tr>
                <td >
                    <span class="wiztxt">This page allows you add or edit static footer content. 
                        </span>
                </td>
            </tr>
            <tr bgcolor="#eeeee0">
                <td align="left" >
                    Document Information</td>
                
            </tr>
            <tr>
                <td>
                 <asp:FormView ID="fvDocument" runat="server" DataSourceID="dsDocument" 
            DefaultMode="Edit" >
           
            <EditItemTemplate>
            <table border="0" width="800">
                   
               
               
                    <tr>
                        <td colspan="3">
                            Content:</td>
                    </tr>
                    <tr style="height:300px" >
                        <td colspan="3" align="left" valign="top" >
                        <FTB:FreeTextBox id="FreeTextBox1" runat="Server" Text='<%# Bind("doc_content") %>' />
                           
                            <asp:RequiredFieldValidator ID="rfvHtext" runat="server" ErrorMessage="Infomation Required"
                                ControlToValidate="FreeTextBox1">
                            </asp:RequiredFieldValidator></td>
                    </tr>
                  
                    <tr>
                        <td align="center" colspan="3">
                        <input type="button" onclick ="JavaScript:window.history.back();" value ="Cancel" />
                        &nbsp;&nbsp;&nbsp;
                            <asp:Button ID="btnUpdateContent" runat="server" CommandName="Update" Text="Update"></asp:Button>
                            </td>
                    </tr>
                </table>
            
            </EditItemTemplate>
        </asp:FormView>
                </td>
            </tr>
        </table>
       
    </asp:Panel>
    
    <asp:SqlDataSource ID="dsDocument" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
        SelectCommand="GetDocumentsByDocID" SelectCommandType="StoredProcedure" UpdateCommand="UpdateDocumentContent" UpdateCommandType="StoredProcedure"
         >
     
        <SelectParameters>
            <asp:QueryStringParameter DefaultValue="0" Name="doc_id" Type="Int32" QueryStringField="doc_id" />
        </SelectParameters>
      
        <UpdateParameters>
            <asp:QueryStringParameter Name="doc_id" QueryStringField="doc_id" DefaultValue="0" Type="Int32"  />
            <asp:Parameter Name="doc_content" Type="String" />
        </UpdateParameters>
    </asp:SqlDataSource>
    
</asp:Content>
