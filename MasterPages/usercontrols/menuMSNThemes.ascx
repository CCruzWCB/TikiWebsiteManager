<%@ Control Language="C#" ClassName="menuMSNThemes" %>
<%@ Implements Interface="System.Web.UI.WebControls.WebParts.IWebPart" %>

<script runat="server">
    private string _catalogiconimageurl;
    private string _description;
    private string _subtitle;
    private string _title;
    private string _titleiconimageurl;
    private string _titleurl;

    string IWebPart.CatalogIconImageUrl
    {
        get { return _catalogiconimageurl; }
        set { _catalogiconimageurl = value; }
    }
    string IWebPart.Description
    {
        get { return _description; }
        set { _description = value; }
    }
    string IWebPart.Subtitle
    {
        get { return _subtitle; }
    }
    string IWebPart.Title
    {
        get { return _title; }
        set { _title = value; }
    }
    string IWebPart.TitleIconImageUrl
    {
        get { return _titleiconimageurl; }
        set { _titleiconimageurl = value; }
    }

    string IWebPart.TitleUrl
    {
        get { return _titleurl; }
        set { _titleurl = value; }
    }

    private void Page_Load(object sender, System.EventArgs e)
    {
        _title = "  MSN Themes";
        _description = "";
        _titleiconimageurl = "~/images/menuicon.jpg";
    }
    

</script>

<br />
<ul class="menutextindent">
    <li><a href="msnblue.aspx">MSN Blue</a></li>
    <li><a href="msncherryblossom.aspx">MSN Cherry Blossom</a></li>
    <li><a href="msnfinance.aspx">MSN Finance</a></li>
    <li><a href="msnmorning.aspx">MSN Morning</a></li>
    <li><a href="msnpurple.aspx">MSN Purple</a></li>
    <li><a href="msnred.aspx">MSN Red</a></li>
</ul>



