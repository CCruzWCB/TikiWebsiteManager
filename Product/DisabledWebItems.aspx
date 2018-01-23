﻿<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="DisabledWebItems.aspx.vb" Inherits="Product_DisabledWebItems" Title="Manage Items with disabled 'Add To Cart' button" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1"  %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
<asp:ScriptManager ID="ScriptManager1" runat="server">
                                </asp:ScriptManager>
<style>



.button {
	-moz-box-shadow:inset 0px 1px 0px 0px #ffffff;
	-webkit-box-shadow:inset 0px 1px 0px 0px #ffffff;
	box-shadow:inset 0px 1px 0px 0px #ffffff;
	background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #ededed), color-stop(1, #dfdfdf) );
	background:-moz-linear-gradient( center top, #ededed 5%, #dfdfdf 100% );
	filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#ededed', endColorstr='#dfdfdf');
	background-color:#ededed;
	-moz-border-radius:6px;
	-webkit-border-radius:6px;
	border-radius:6px;
	border:1px solid #dcdcdc;
	display:inline-block;
	color:#777777;
	font-family:arial;
	font-size:15px;
	font-weight:bold;
	padding:0px 6px;
	text-decoration:none;
	text-shadow:1px 1px 0px #ffffff;
}.button:hover {
	background:-webkit-gradient( linear, left top, left bottom, color-stop(0.05, #dfdfdf), color-stop(1, #ededed) );
	background:-moz-linear-gradient( center top, #dfdfdf 5%, #ededed 100% );
	filter:progid:DXImageTransform.Microsoft.gradient(startColorstr='#dfdfdf', endColorstr='#ededed');
	background-color:#dfdfdf;
	 cursor: pointer;
}.button:active {
	position:relative;
	top:1px;
}
/* This imageless css button was generated by CSSButtonGenerator.com */
.tbg {	font: bold 11px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;	color: #4f6b72;	border-right: 1px solid #C1DAD7;	border-bottom: 1px solid #C1DAD7;	border-top: 1px solid #C1DAD7;	letter-spacing: 2px;	text-transform: uppercase;	text-align: left;	padding: 6px 6px 6px 12px;	background: #CAE8EA url(../images/bg_header.jpg) no-repeat;}#mytable {	width: 700px;	padding: 0;	margin: 0;}caption {	padding: 0 0 5px 0;	width: 700px;	 	font: italic 11px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;	text-align: right;}th {	font: bold 11px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;	color: #4f6b72;	border-right: 1px solid #C1DAD7;	border-bottom: 1px solid #C1DAD7;	border-top: 1px solid #C1DAD7;	letter-spacing: 2px;	text-transform: uppercase;	text-align: left;	padding: 6px 6px 6px 12px;	background: #CAE8EA url(../images/bg_header.jpg) no-repeat;}th.nobg {	border-top: 0;	border-left: 0;	border-right: 1px solid #C1DAD7;	background: none;}td {	border-right: 1px solid #C1DAD7;	border-left: 1px solid #C1DAD7;	border-bottom: 1px solid #C1DAD7;	background: #fff;	padding: 6px 6px 6px 12px;	color: #4f6b72;}td.alt {	background: #F5FAFA;	color: #797268;}th.spec {	border-left: 1px solid #C1DAD7;	border-top: 0;	background: #fff url(images/bullet1.gif) no-repeat;	font: bold 10px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;}th.specalt {	border-left: 1px solid #C1DAD7;	border-top: 0;	background: #f5fafa url(images/bullet2.gif) no-repeat;	font: bold 10px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;	color: #797268;}
 h4 
          {
              padding-bottom:5px;
            }
            
            .tag {font-weight:bold;padding-bottom:2px;}
            .value {padding-bottom:2px;padding-left:4px;color:#4097ca;font-size:14px;}
            
              
</style>

<asp:Panel ID="pnlAdd" runat="server" DefaultButton="btnAdd">
<div class="tbg" style="width:480px;text-align:right;">
Disable Add To Cart for Item &nbsp; <asp:TextBox ID="txtItem" runat="server"></asp:TextBox> &nbsp; <asp:Button ID="btnAdd" CssClass="button" runat="server" Text="add" />
<cc1:AutoCompleteExtender ID="acProductModelNumber" runat="server" SkinID="ModelAutoComplete"  TargetControlID ="txtItem" > </cc1:AutoCompleteExtender>

</div>
</asp:Panel>
 <asp:GridView ID="gvDisabledAddToCartItems" style="width:500px" runat="server" AutoGenerateColumns="False" DataKeyNames="ProductID"  DataSourceID="dsDisableItems">
    
    <Columns>
                                                                        
      <asp:TemplateField>
        <ItemTemplate>
            <asp:Image ID="img" runat="server" ImageUrl='<%#  utilities.getimageURL(eval("ImagePath").tostring) %>' />
        </ItemTemplate>
      </asp:TemplateField>
        <%--<asp:BoundField DataField="RefreshItem" HeaderText="Item" SortExpression="RefreshType" />--%>
        <asp:BoundField DataField="ProductModelNumber" HeaderText="Item Number" SortExpression="ProductModelNumber" />
        <asp:BoundField DataField="Name" HeaderText="Name" SortExpression="Name" />
<%--        <asp:BoundField DataField="Active" HeaderText="Active"  SortExpression="Active" />
        <asp:BoundField DataField="IsSellable" HeaderText="IsSellable"  SortExpression="IsSellable" />
--%>        <asp:CommandField  ShowDeleteButton="true"  HeaderText="Action" ControlStyle-ForeColor="#2c71b6" ControlStyle-Font-Bold="true" DeleteText="remove"    />
                                                                    
    </Columns>
</asp:GridView>
<asp:SqlDataSource ID="dsDisableItems" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
    SelectCommand="[GetDisableAddToCartItems]" SelectCommandType="StoredProcedure" 
    UpdateCommand="[SetDisableAddToCartFlag]" UpdateCommandType="StoredProcedure" 
    DeleteCommand="[SetDisableAddToCartFlag]" DeleteCommandType="StoredProcedure" 
    >
    
    <SelectParameters>
        <asp:Parameter Name="BrandID" Type="string" DefaultValue="<%$ AppSettings:BrandID %>" />
    </SelectParameters>

  
    <UpdateParameters>
        <asp:Parameter Name="ProductID" Type="Int32" DefaultValue="0" />
        <asp:Parameter Name="ProductModelNumber" Type="String"  />
        <asp:Parameter Name="IsDisableAddToCart" Type="Boolean" DefaultValue="True"  />
    </UpdateParameters>
    
    <DeleteParameters>
        <asp:Parameter Name="ProductID" Type="Int32" DefaultValue="0" />
        <asp:Parameter Name="IsDisableAddToCart" Type="Boolean" DefaultValue="False"  />
    </DeleteParameters>
                                                                    
    </asp:SqlDataSource>



</asp:Content>

