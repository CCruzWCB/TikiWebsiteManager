<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="SearchRedirects.aspx.vb" Inherits="Media_SearchRedirects" title="Manage Search Redirects"     %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">


<div style="width:70%;margin:5px;" class ="box">
<div style="margin:5px;" class ="graybox" ><b>Add Search Redirect </b></div>

<asp:Panel ID="pnlAdd" runat="server" DefaultButton="btnAddNew"  >
<table  class ="graybox" >
<tr><td align ="left" >Search Term:<asp:TextBox ID="txtName"  runat="server" width="150"  MaxLength ="100" ></asp:TextBox>
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
<div style="margin:5px;" class ="graybox" ><b>Manage Search Redirects <asp:Label ID="lblCount" runat="server" Text ="(0)"></asp:Label></b></div>
<asp:GridView ID="gvSearchRedirect" DataKeyNames="SearchRedirectID" runat="server" DataSourceID="dsSearchRedirect" 
        AutoGenerateColumns="False">
        <Columns>
            <asp:BoundField DataField="SearchRedirectID" HeaderText="ID" 
                InsertVisible="False"  ReadOnly="true" SortExpression="SearchRedirectID"  />
            <asp:BoundField DataField="Name" ItemStyle-HorizontalAlign="left"  ItemStyle-Width="100" ControlStyle-Width="100"  
                HeaderText="Name" SortExpression="Name" HeaderStyle-HorizontalAlign="left" >
<HeaderStyle HorizontalAlign="Left"></HeaderStyle>

<ItemStyle HorizontalAlign="Left"></ItemStyle>
            </asp:BoundField>
            <asp:BoundField DataField="RedirectURL" ItemStyle-Width="400" ControlStyle-Width="400" HeaderText="RedirectURL" 
                ItemStyle-HorizontalAlign="left" HeaderStyle-HorizontalAlign="left"
                SortExpression="RedirectURL" >
<HeaderStyle HorizontalAlign="Left"></HeaderStyle>

<ItemStyle HorizontalAlign="Left"></ItemStyle>
            </asp:BoundField>
            <asp:BoundField DataField="DateCreated" HeaderText="DateCreated" 
                SortExpression="DateCreated" ReadOnly="True" />
            <asp:CheckBoxField DataField="Active" HeaderText="Active" 
                SortExpression="Active" />
            <asp:CommandField InsertVisible="False" ShowDeleteButton="True" 
                ShowEditButton="True" />
        </Columns>
    </asp:GridView>
</div> 

<asp:SqlDataSource ID="dsSearchRedirect" runat="server" 
        ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>" 
        DeleteCommand="Manager.DeleteSearchRedirect" DeleteCommandType="StoredProcedure" 
        InsertCommand="Manager.AddSearchRecirect" InsertCommandType="StoredProcedure" 
        SelectCommand="Manager.GetALLSearchRedirect" SelectCommandType="StoredProcedure" 
        UpdateCommand="Manager.UpdateSearchRedirect" UpdateCommandType="StoredProcedure">
        <DeleteParameters>
            <asp:Parameter Name="SearchRedirectID" Type="Int32" />
        </DeleteParameters>
        <UpdateParameters>
            <asp:Parameter Name="SearchRedirectID" Type="Int32"  />
            <asp:Parameter Name="Name" Type="String" Size="100" />
            <asp:Parameter Name="RedirectURL" Type="String" Size="250" />
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

