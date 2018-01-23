<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="AddEditPromotion.aspx.vb" Inherits="Promotion_AddPromotion" title="Untitled Page" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
    <asp:FormView ID="fvPromotion" runat="server" DataKeyNames="PromotionID" DataSourceID="dsPromotion" DefaultMode="insert">
          <EditItemTemplate>
         <table>
         <tr><td>Name:</td><td><asp:TextBox ID="NameTextBox" runat="server" Text='<%# Bind("Name") %>'>
            </asp:TextBox></td></tr>
            <tr><td>Description:</td><td>
            <asp:TextBox ID="DescriptionTextBox" runat="server" Text='<%# Bind("Description") %>'>
            </asp:TextBox></td></tr> 
            <tr><td>StartDate:</td><td>
            <asp:TextBox ID="StartDateTextBox" runat="server" Text='<%# Bind("StartDate") %>'>
            </asp:TextBox></td></tr> 
            <tr><td>EndDate:</td><td>
            <asp:TextBox ID="EndDateTextBox" runat="server" Text='<%# Bind("EndDate") %>'>
            </asp:TextBox></td></tr> 
            <tr><td>Active:</td><td>
             <asp:CheckBox ID="ActiveCheckBox" runat="server" Checked='<%# Bind("Active") %>' /></td></tr>
             <tr><td colspan ="2" align ="center" >
            <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" CommandName="Update"
                Text="Update">
            </asp:LinkButton>&nbsp;&nbsp;
           <asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" PostBackUrl ="~/Promotion/ManagePromotions.aspx"  
                Text="Cancel">
            </asp:LinkButton></td></tr> 
            </table>
        </EditItemTemplate>
        
        <InsertItemTemplate>
        <table>
            <tr><td>Name:</td><td>
            <asp:TextBox ID="NameTextBox" runat="server" Text='<%# Bind("Name") %>'>
            </asp:TextBox></td></tr> 
            <tr><td>Description:</td><td>
            <asp:TextBox ID="DescriptionTextBox" runat="server" Text='<%# Bind("Description") %>'>
            </asp:TextBox></td></tr> 
            <tr><td>StartDate:</td><td>
            <asp:TextBox ID="StartDateTextBox" runat="server" Text='<%# Bind("StartDate") %>'>
            </asp:TextBox></td></tr> 
            <tr><td>EndDate:</td><td>
            <asp:TextBox ID="EndDateTextBox" runat="server" Text='<%# Bind("EndDate") %>'>
            </asp:TextBox></td></tr> 
            <tr><td>Active:</td><td>
            <asp:CheckBox ID="ActiveCheckBox" runat="server" Checked='<%# Bind("Active") %>' /></td></tr> 
            <tr><td colspan ="2" align ="center" ><asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" CommandName="Insert"
                Text="Insert">
            </asp:LinkButton>&nbsp;&nbsp;
            <asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" PostBackUrl ="~/Promotion/ManagePromotions.aspx"  
                Text="Cancel">
            </asp:LinkButton></td></tr> 
            </table>
        </InsertItemTemplate>
        
    </asp:FormView>
    <asp:SqlDataSource ID="dsPromotion" runat="server" ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>"
        InsertCommand="AddPromotion" InsertCommandType="StoredProcedure" UpdateCommand="UpdatePromotion" UpdateCommandType="StoredProcedure"  SelectCommand="GetPromotion" SelectCommandType="StoredProcedure"  >
       <SelectParameters>
            <asp:QueryStringParameter Name="PromotionID" QueryStringField="PID" DefaultValue="0" />
       </SelectParameters>
        <UpdateParameters>
            <asp:Parameter Name="Name" Type="String" />
            <asp:Parameter Name="Description" Type="String" />
            <asp:Parameter Name="StartDate" Type="DateTime" />
            <asp:Parameter Name="EndDate" Type="DateTime" />
            <asp:Parameter Name="Active" Type="Boolean" />
            <asp:Parameter Name="PromotionID" Type="Int32" />
        </UpdateParameters>
        <InsertParameters>
            <asp:Parameter Name="Name" Type="String" />
            <asp:Parameter Name="Description" Type="String" />
            <asp:Parameter Name="StartDate" Type="DateTime" />
            <asp:Parameter Name="EndDate" Type="DateTime" />
            <asp:Parameter Name="Active" Type="Boolean" />
            <asp:Parameter Name="PromotionID" Type="int32" Direction="InputOutput"  />
        </InsertParameters>
    </asp:SqlDataSource>


</asp:Content>

