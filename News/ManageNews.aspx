<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="ManageNews.aspx.vb" Inherits="Media_ManageNews" title="Manage News" ValidateRequest ="false"  %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">

<div style="width:70%;margin:5px;" class ="box">

<div class ="graybox" >
<div style="height:20px;"><b>Manage News <asp:Label ID="lblCount" runat="server" Text ="(0)"></asp:Label></b>
<span style="height:20px;left:215px;top:0px;position:relative ;"><a href="AddNews.aspx" title ="Click to add a new News item" >Click to Add a News Items</a></span>
</div>

</div>

<br /><br />
<asp:GridView ID="gvNews" runat="server" AllowPaging="True" AllowSorting="True"  Width ="100%"
        AutoGenerateColumns="False" DataKeyNames="KnowledgeID"  PageSize ="25" 
        DataSourceID="dsGetNews" >
    <Columns>
        <asp:TemplateField HeaderText="KnowledgeID" InsertVisible="False"  Visible ="false" 
            SortExpression="KnowledgeID">
            <ItemTemplate>
                <asp:Label ID="Label1" runat="server" Text='<%# Bind("KnowledgeID") %>'></asp:Label>
                <asp:Label ID="lblFilePath"  Visible="false" runat="server" Text='<%# Eval("FilePath") %>'></asp:Label>
            </ItemTemplate>
        </asp:TemplateField>
        <asp:TemplateField HeaderText="Title" SortExpression="Title" ItemStyle-HorizontalAlign ="Left" >
            <ItemTemplate>
            <a href ='<%# "EditNews.aspx?KnowledgeID=" & Eval("KnowledgeID") %>' title ='Click to Edit News Article items' ><%#Eval("Title")%></a>
            </ItemTemplate>
        </asp:TemplateField>
        <asp:TemplateField HeaderText="Category" SortExpression="CategoryName">
            <ItemTemplate>
                <asp:Label ID="lblCategory" runat="server" Text='<%# Bind("CategoryName") %>'></asp:Label>
            </ItemTemplate>
        </asp:TemplateField>
        <asp:TemplateField HeaderText="Active" SortExpression="Active">
            <ItemTemplate>
                <asp:Label ID="lblyes" Text ="Yes" runat="server" Visible='<%# Cbool(Eval("Active")) %>' ></asp:Label>
                    <asp:Label ID="lblNo" Text ="No" runat="server" Visible='<%# Not Cbool(Eval("Active")) %>'></asp:Label>
            </ItemTemplate>
        </asp:TemplateField>
        
                <asp:TemplateField HeaderText="Date Created" SortExpression="DateCreated">
            <ItemTemplate>
                <asp:Label ID="Label2" runat="server" Text='<%# Bind("DateCreated") %>'></asp:Label>
            </ItemTemplate>
        </asp:TemplateField>

        <asp:TemplateField ShowHeader="False">
            <ItemTemplate>
                <asp:LinkButton ID="lnkDelete" runat="server" OnClientClick ="return window.confirm('Delete Current News Item. Are You Sure?');" CausesValidation="False" 
                    CommandName="Delete" Text="Delete"></asp:LinkButton>
            </ItemTemplate>
        </asp:TemplateField>
    </Columns>
</asp:GridView>

    <asp:SqlDataSource ID="dsGetNews" runat="server" 
        ConnectionString="<%$ ConnectionStrings:LLFPublicWebsiteConnectionString %>" 
        SelectCommand="llfbps.dbo.ManagerGetWebsiteNews" SelectCommandType="StoredProcedure" DeleteCommand="llfbps.dbo.DeleteKnowledge" 
        DeleteCommandType="StoredProcedure">
        <DeleteParameters>
            <asp:Parameter Name="KnowledgeID" Type="Int32" />
        </DeleteParameters>
    </asp:SqlDataSource>
</div> 
</asp:Content>

