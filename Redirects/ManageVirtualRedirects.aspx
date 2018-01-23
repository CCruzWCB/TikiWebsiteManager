<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="ManageVirtualRedirects.aspx.vb" Inherits="Features_ManageVirtualRedirects" title="Manage URL Redirects" CodeFileBaseClass="BasePage"   %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">


<div style="width:70%;margin:5px;" class ="box">
<div style="margin:5px;" class ="graybox" ><b>Add URL Folder Redirect </b></div>

<asp:Panel ID="pnlAdd" runat="server" DefaultButton="btnAddNew">
<table  class ="graybox" >
<tr><td align ="left" >Virtual Folder Name:<asp:TextBox ID="txtName"  runat="server" width="150"  MaxLength ="100" ></asp:TextBox>
<asp:RequiredFieldValidator ID="rfvName" runat="server" ValidationGroup="AddSearchRedirect" ControlToValidate ="txtName" Display ="Static" ErrorMessage ="Name is required" >&nbsp;</asp:RequiredFieldValidator></td>
    <td align ="left" >Redirect URL:<asp:TextBox ID="txtURL" runat="server" width="300" MaxLength ="250" ></asp:TextBox>
    <asp:RequiredFieldValidator ID="rfvURL" runat="server"  ValidationGroup="AddSearchRedirect"  Display="Static"  ControlToValidate ="txtURL" ErrorMessage ="URL is required" >&nbsp;</asp:RequiredFieldValidator>
    </td>
<td align ="center" >Active: <asp:CheckBox ID="chkActive" runat="server" /></td>    
<td><asp:Button ID="btnAddNew" runat="server" ValidationGroup="AddSearchRedirect"  Text ="Add New" /></td>
    </tr>
    <tr>
        <td colspan="3">
            <asp:ValidationSummary ID="vsSummary" ValidationGroup="AddSearchRedirect" runat="server" /> </td>
    </tr>

</table>
</asp:Panel>

</div>

<br /><br />

<div style="width:70%;margin:5px;" class ="box">
<div style="margin:5px;" class ="graybox" ><b>Manage URL Virtual Folder Redirects <asp:Label ID="lblCount" runat="server" Text ="(0)"></asp:Label></b></div>
<asp:GridView ID="gvVirtualRedirect" DataKeyNames="RedirectID" runat="server" DataSourceID="dsVirtualRedirect" 
        AutoGenerateColumns="False">
        <Columns>
            <asp:BoundField DataField="RedirectID" HeaderText="ID" 
                InsertVisible="False"  ReadOnly="true" SortExpression="SearchRedirectID"  />
            <asp:BoundField DataField="Name" ItemStyle-HorizontalAlign="left"  ItemStyle-Width="100" ControlStyle-Width="100"  
                HeaderText="Folder Name" SortExpression="Name" HeaderStyle-HorizontalAlign="left" >
<HeaderStyle HorizontalAlign="Left"></HeaderStyle>

<ItemStyle HorizontalAlign="Left"></ItemStyle>
            </asp:BoundField>
            <asp:BoundField DataField="RedirectURL" ItemStyle-Width="400" ControlStyle-Width="400" HeaderText="RedirectURL" 
                ItemStyle-HorizontalAlign="left" HeaderStyle-HorizontalAlign="left"
                SortExpression="RedirectURL" >
<HeaderStyle HorizontalAlign="Left"></HeaderStyle>

<ItemStyle HorizontalAlign="Left"></ItemStyle>
            </asp:BoundField>
            <asp:BoundField DataField="CreatedByName" HeaderText="Created By" 
                SortExpression="CreatedByName" ReadOnly="True" />
            <asp:CheckBoxField DataField="Active" HeaderText="Active" 
                SortExpression="Active" />
            <asp:CommandField InsertVisible="False" ShowDeleteButton="True" 
                ShowEditButton="True" />
        </Columns>
    </asp:GridView>
</div> 

<asp:SqlDataSource ID="dsVirtualRedirect" runat="server" 
        ConnectionString="<%$ ConnectionStrings:DBConnectionWebsite %>" 
        DeleteCommand="Manager.DeleteVirtualFolderRedirect" DeleteCommandType="StoredProcedure" 
        InsertCommand="Manager.AddVirtualFolderRedirect" InsertCommandType="StoredProcedure" 
        SelectCommand="Manager.GetAllVirtualFolderRedirects" SelectCommandType="StoredProcedure" 
        UpdateCommand="Manager.UpdateVirtualFolderRedirect" UpdateCommandType="StoredProcedure">
        <DeleteParameters>
            <asp:Parameter Name="RedirectID" Type="Int32" />
        </DeleteParameters>
        <UpdateParameters>
            <asp:Parameter Name="RedirectID" Type="Int32"  />
            <asp:Parameter Name="Name" Type="String" Size="100" />
            <asp:Parameter Name="RedirectURL" Type="String" Size="250" />
            <asp:Parameter Name="LastUpdatedBy" Type="Int32"  />
            <asp:Parameter Name="Active" Type="Boolean" />
        </UpdateParameters>
        <InsertParameters>
            <asp:Parameter Name="Name" Type="String" Size="100" />
            <asp:Parameter Name="RedirectURL" Type="String" Size="250" />
            <asp:Parameter Name="CreatedBy" Type="Int32" />
            <asp:Parameter Name="Active" Type="Boolean" />
            <asp:Parameter Direction="InputOutput" Name="ReturnValue" Type="Int32" />
        </InsertParameters>
    </asp:SqlDataSource>


   
        










    
    
</asp:Content>

