<%@ Page Language="C#" MasterPageFile="~/MasterPages/Default.master" Title="Untitled Page" StylesheetTheme="MSN_Blue"%>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentPlaceHolder1" Runat="Server">
    <br />
     <p class="title">
    MSN Blue Theme
        </p>
        <hr>
   <br />
    <h1>
        Heading 1
    </h1>
    <p class="bodytextindent1">
    Text
    </p>

    <h2>
        Heading 2<br />
    </h2>
    <p class="bodytextindent1">
    Text
    </p>
    <h3>
        Heading 3</h3>
    <p class="bodytextindent1">
    Text
    </p>
    <h3>
        Sample DetailsView</h3>
        <br />
    <asp:DetailsView ID="DetailsView1" runat="server" SkinID="DetailsView" DataSourceID="SiteMapDataSource1">
        <Fields>
            <asp:BoundField DataField="Description" HeaderText="Description" ReadOnly="True"
                SortExpression="Description" />
            <asp:BoundField DataField="Title" HeaderText="Title" ReadOnly="True" SortExpression="Title" />
            <asp:BoundField DataField="Url" HeaderText="Url" ReadOnly="True" SortExpression="Url" />
        </Fields>
    </asp:DetailsView>
        <br />
    <h3>
        Sample GridView</h3>
        <br />
    <asp:GridView ID="GridView1" runat="server" SkinID="GridView" DataSourceID="SiteMapDataSource1">
        <Columns>
            <asp:BoundField DataField="Title" HeaderText="Title" ReadOnly="True" SortExpression="Title" />
            <asp:BoundField DataField="Description" HeaderText="Description" ReadOnly="True" SortExpression="Description" />
            <asp:BoundField DataField="Url" HeaderText="Url" ReadOnly="True" SortExpression="Url" />
        </Columns>
    </asp:GridView>
    <asp:SiteMapDataSource ID="SiteMapDataSource1" runat="server" StartFromCurrentNode="true" />
     <br />
     <h3>
        Sample Calendar</h3>
    <br />
    <asp:Calendar ID="Calendar1" runat="server" SkinID="CalendarView"/>
    <br />
    <h3>Sample FormView Control</h3>
    <br />
    <asp:FormView ID="FormView1" runat="server" SkinID="FormView" DataSourceID="SiteMapDataSource1">
        <EditItemTemplate>
            Description:
            <asp:TextBox ID="DescriptionTextBox" runat="server" Text='<%# Bind("Description") %>'>
            </asp:TextBox><br />
            Title:
            <asp:TextBox ID="TitleTextBox" runat="server" Text='<%# Bind("Title") %>'>
            </asp:TextBox><br />
            Url:
            <asp:TextBox ID="UrlTextBox" runat="server" Text='<%# Bind("Url") %>'>
            </asp:TextBox><br />
            <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="True" CommandName="Update"
                Text="Update">
            </asp:LinkButton>
            <asp:LinkButton ID="UpdateCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                Text="Cancel">
            </asp:LinkButton>
        </EditItemTemplate>
        <InsertItemTemplate>
            Description:
            <asp:TextBox ID="DescriptionTextBox" runat="server" Text='<%# Bind("Description") %>'>
            </asp:TextBox><br />
            Title:
            <asp:TextBox ID="TitleTextBox" runat="server" Text='<%# Bind("Title") %>'>
            </asp:TextBox><br />
            Url:
            <asp:TextBox ID="UrlTextBox" runat="server" Text='<%# Bind("Url") %>'>
            </asp:TextBox><br />
            <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" CommandName="Insert"
                Text="Insert">
            </asp:LinkButton>
            <asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                Text="Cancel">
            </asp:LinkButton>
        </InsertItemTemplate>
        <ItemTemplate>
            Description:
            <asp:Label ID="DescriptionLabel" runat="server" Text='<%# Bind("Description") %>'></asp:Label><br />
            Title:
            <asp:Label ID="TitleLabel" runat="server" Text='<%# Bind("Title") %>'></asp:Label><br />
            Url:
            <asp:Label ID="UrlLabel" runat="server" Text='<%# Bind("Url") %>'></asp:Label><br />
        </ItemTemplate>
    </asp:FormView>
    <br />
     <h3>
        Sample Login Control</h3>
    <br />
    <asp:Login runat=server ID="Login1" SkinID="LoginView" />
    <br />

</asp:Content>

