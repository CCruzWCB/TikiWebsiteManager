<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" ValidateRequest ="false"  AutoEventWireup="false" CodeFile="EditNews.aspx.vb" Inherits="Media_EditNews" title="Edit News Item" %>
<%@ Register TagPrefix="FTB" Namespace="FreeTextBoxControls" Assembly="FreeTextBox" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">

<div style="width:70%;margin:5px;" class ="box">
<div style="margin:5px;" class ="graybox" ><b>Edit News Item: <asp:Label ID="lblTitle" runat="server" ></asp:Label></b></div>

<br />
<br />

<table border="0" width="850px"  class="box" >
<tr ><td align ="right" >Title:</td><td align ="left" ><asp:TextBox ID="txtTitle" runat="server" Columns ="70"  MaxLength ="250" ></asp:TextBox>&nbsp;&nbsp;&nbsp;
<asp:RequiredFieldValidator ID="rfvTitle" runat="server" ControlToValidate ="txtTitle" ErrorMessage ="Title is required" >&nbsp;</asp:RequiredFieldValidator></td></tr>
<tr><td align ="right"> Category: </td><td align ="left" ><asp:DropDownList ID="ddCategory" runat="server" >
<asp:ListItem Value ="9" Text ="General News"></asp:ListItem>

</asp:DropDownList></td></tr>
<tr><td align ="right"> Active: </td><td align ="left" ><asp:CheckBox ID="chkActive" runat="server"  Text ="Show on Website"/></td></tr>
<tr><td align ="right" >News Source: </td>
<td align ="left" ><asp:RadioButtonList ID="rdoContentType" runat="server" AutoPostBack ="true"  RepeatDirection ="Horizontal">
<asp:ListItem Value ="1" Selected ="true" Text ="Text/HTML" ></asp:ListItem>
<asp:ListItem Value ="2" Text ="URL" ></asp:ListItem>
<asp:ListItem Value ="3" Text ="File (PDF)" ></asp:ListItem>
</asp:RadioButtonList>
</td></tr>
<tr><td colspan="2" width="800px" align="center" ><br /><br />
<asp:Panel ID="pHtmlTExt" CssClass ="graybox" runat="server" Width ="800px" Visible ="True">
<FTB:FreeTextBox ID="txtArticle" 
ToolbarLayout="ParagraphMenu, FontFacesMenu, FontSizesMenu, FontForeColorsMenu, FontForeColorPicker, FontBackColorsMenu|Bold, Italic, Underline, Strikethrough, Superscript, Subscript|JustifyLeft, JustifyRight, JustifyCenter, JustifyFull|BulletedList, NumberedList, Indent, Outdent|Cut, Copy, Paste, Delete, Undo, Redo|SymbolsMenu, InsertRule, InsertDate, InsertTime|InsertTable;InsertTableRowBefore, InsertTableRowAfter, DeleteTableRow;InsertTableColumnBefore, InsertTableColumnAfter, DeleteTableColumn;EditTable|InsertDiv;InsertImageFromGallery, InsertImage;Preview;SelectAll, EditStyle, WordClean" 
ImageGalleryPath ="~/BPS/Lamplight/Lamplight/MKTGContent" 

Width ="750px" 
Height ="650px" 
runat="Server" 
/>
<asp:RequiredFieldValidator ID="rfvArticle" runat="server"  Enabled ="false" ControlToValidate ="txtArticle" ErrorMessage ="HTML/Text Content is required" >&nbsp;</asp:RequiredFieldValidator>
</asp:Panel>

<asp:Panel ID="pURL" CssClass ="graybox"  runat="server" Visible ="false" Width ="800px">
Sample URL format ("http://www.mywebsite.com/targetpagename.html")<br />
<asp:TextBox ID="txtURL" Columns ="85" MaxLength ="500" runat="server"></asp:TextBox>
<asp:RequiredFieldValidator ID="rfvURL" runat="server" Enabled ="false" ControlToValidate ="txtURL" ErrorMessage ="Please enter a URL" >&nbsp;</asp:RequiredFieldValidator>
<br />
URL Target:<asp:DropDownList ID="ddTargetSource" runat="server" >
<asp:ListItem Value ="0" Text ="Internet URL (not on Lamplight.com site)"></asp:ListItem>
<asp:ListItem Value ="1" Text ="URL existing on Lamplight.com"></asp:ListItem>
</asp:DropDownList>
</asp:Panel>

<asp:Panel ID="pUpload"  runat="Server" CssClass ="graybox"  Visible ="false" Width ="800px">
PDF file:<asp:FileUpload ID="fileUpload" runat="server" />&nbsp;&nbsp;&nbsp;
<asp:HyperLink ID="hplCurrentFile"  Target ="_blank" runat="server" NavigateUrl ="" Text ="View Current File" Visible ="false" ></asp:HyperLink>
<asp:Label id="lblCurrentFilePath" runat="server" Visible ="false" ></asp:Label>
<asp:RequiredFieldValidator ID="rfvUpload" Enabled ="false" runat="server" ControlToValidate ="fileUpload" ErrorMessage ="Please select a file" >&nbsp;</asp:RequiredFieldValidator>
</asp:Panel>
<br /><br />
<asp:Button ID="btnCancel" runat="server"  Text ="Cancel"  CausesValidation ="false" PostBackUrl ="~/News/ManageNews.aspx" />&nbsp;&nbsp;&nbsp;
<asp:Button ID="btnEdit" runat="server"  Text ="Save Changes" />
</td></tr>
</table> 
</div>


</asp:Content>

