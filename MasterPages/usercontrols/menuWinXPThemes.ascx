<%@ Control Language="C#" ClassName="menuWinXPThemes" %>
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
        _title = "  WIN XP Themes";
        _description = "";
        _titleiconimageurl = "~/images/menuicon.jpg";
    }
    

</script>
<br />
<ul class="menutextindent">
    <li><a href="winxpblue.aspx">WIN XP Blue</a></li>
    <li><a href="winxpsilver.aspx">WIN XP Silver</a></li>
</ul>