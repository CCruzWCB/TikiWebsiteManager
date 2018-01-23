<%@ Page Language="VB" AutoEventWireup="false" MasterPageFile="~/MasterPages/MasterPage.master" CodeFile="ManageProduct.aspx.vb" Inherits="Manager_ManageProduct"
    ValidateRequest="false" %>

<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>

<%@ Register TagPrefix="FTB" Namespace="FreeTextBoxControls" Assembly="FreeTextBox" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<asp:Content ContentPlaceHolderID="contentMain" ID="ContentMain" runat="server">


        <script type="text/javascript" language="javascript">
function onUpdating(){
        // get the update progress div
        
        //alert ('here');
        var updateProgressDiv = $get('updateProgressDiv'); 

        //  get the gridview element        
        var gridView = $get('<%= pnlForm.ClientID %>');
        //var gridView = document.forms[0]['pnlForm'];

        //alert (gridView);
        // make it visible
        updateProgressDiv.style.display = '';        
        
        // get the bounds of both the gridview and the progress div
        var gridViewBounds = Sys.UI.DomElement.getBounds(gridView);
        var updateProgressDivBounds = Sys.UI.DomElement.getBounds(updateProgressDiv);
        
        //  center of gridview
        var x = gridViewBounds.x + Math.round(gridViewBounds.width / 2) - Math.round(updateProgressDivBounds.width / 2);
        var y = gridViewBounds.y + Math.round(gridViewBounds.height / 2) - Math.round(updateProgressDivBounds.height / 2);        

        //    set the progress element to this position
        Sys.UI.DomElement.setLocation (updateProgressDiv, x, y);           
    }

    function onUpdated() {
        // get the update progress div
        var updateProgressDiv = $get('updateProgressDiv'); 
        // make it invisible
        updateProgressDiv.style.display = 'none';
    }

        </script>

        <asp:ScriptManager ID="smMain" runat="server" EnablePartialRendering="true">
        </asp:ScriptManager>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender2" runat="server"
            TargetControlID="TabContainer$tpProductInfo$UpdatePanelProduct">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender4" runat="server"
            TargetControlID="TabContainer$tpParts$UpdatePanelBOM">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender5" runat="server"
            TargetControlID="TabContainer$tpProductDiagrams$UpdatePanelDiagramMain">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender6" runat="server"
            TargetControlID="TabContainer$tpProductImages$UpdatePanelImages">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender7" runat="server"
            TargetControlID="TabContainer$tpProductSpecs$UpdatePanelSpecs">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender8" runat="server"
            TargetControlID="TabContainer$tpProductManuals$UpdatePanelProductManuals">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender1" runat="server"
            TargetControlID="TabContainer$tpProductRetailers$UpdatePanelRetailers">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender9" runat="server"
            TargetControlID="TabContainer$tpCrossells$UpdatePanelCrossells">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender11" runat="server"
            TargetControlID="TabContainer$tpProductDescription$upProductDescription">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender3" runat="server"
            TargetControlID="TabContainer$tpProductPromotion$UpdatePanelProductPromotion">
        </cc1:UpdatePanelAnimationExtender>
        
        <asp:Panel ID="pnlForm" runat="server">
            <!-- Page Header Title -->
            <table border="0" cellpadding="3" cellspacing="0" width="100%">
                <tr>
                    <td class="pageheader">
                        &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; <a href="SelectProduct.aspx"
                            target="_self">Select Another Product To Manage</a></td>
                </tr>
            </table>
            <!-- MAIN Table Frame -->
            <table style="width: 100%" border="0" cellpadding="5" cellspacing="0">
                <!-- Error Message -->
                <tr>
                    <td>
                        <asp:UpdatePanel ID="UpdatePanelErrors" runat="server" UpdateMode="Conditional">
                            <ContentTemplate>
                                <table width ="100%" border ="0" >
                                    <tr>
                                        <td width ="100%" >
                                            <asp:Panel ID="pnlError" runat="server" Visible="false"  height ="40px">
                                                <table width="100%" height ="15px" border="0" cellpadding="5" cellspacing="0">
                                                    <tr>
                                                        <td width="35" class="errormsgleft">
                                                            <img width ="20" height ="20" src="../images/msgicon_error.gif" alt="Error" /></td>
                                                        <td class="errormsgright">
                                                            <asp:Label ID="lblError" runat="server" Text="nothing to show"></asp:Label></td>
                                                    </tr>
                                                </table>
                                            </asp:Panel>
                                        </td>
                                    </tr>
                                    <!-- Return Message -->
                                    <tr>
                                        <td>
                                            <asp:Panel ID="pnlReturnMsg" runat="server" Visible="false"  height ="40px">
                                                <table width="100%"  border="0" cellpadding="5" cellspacing="0">
                                                    <tr>
                                                        <td width="35" class="infomsgleft">
                                                            <img width ="20" height ="20" src="../images/msgicon_info.gif" alt="Info Msg" /></td>
                                                        <td  class="infomsgright">
                                                            <asp:Label ID="lblReturnMsg" runat="server"></asp:Label></td>
                                                    </tr>
                                                </table>
                                            </asp:Panel>
                                        </td>
                                    </tr>
                                </table>
                            </ContentTemplate>
                        </asp:UpdatePanel>
                    </td>
                </tr>
                <tr>
                    <td style="width: 100%" align="center" valign="middle">
                        <!--Start Form Here-->
                        <table id="maintable" style="width: 100%" border="0" cellpadding="0" cellspacing="0">
                            <tr>
                                <td class="rcb_topleft">
                                    <img alt="spacer" src="../images/spacer.gif" /></td>
                                <td class="rcb_top">
                                </td>
                                <td class="rcb_topright">
                                    <img alt="spacer" src="../images/spacer.gif" /></td>
                            </tr>
                            
                                        <tr>
                                            <td style="width: 100%" align="left" valign="middle" class="irowbold">
                                            </td>
                                        </tr>
                                    </table>
                                    <!-- END Product Header Content -->
                                </td>
                                <td class="rcb_right">
                                </td>
                            </tr>
                            <tr>
                                <td class="rcb_bottomleft">
                                    <img alt="" src="../images/spacer.gif" /></td>
                                <td class="rcb_bottom">
                                </td>
                                <td class="rcb_bottomright">
                                    <img alt="" src="../images/spacer.gif" /></td>
                            </tr>
                        </table>
                        <table width="100%" border="0" cellpadding="2" cellspacing="4">
                            <tr>
                                <td>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <!-- Message -->
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <!-- Start Tabs -->
                                    <table border="0" cellpadding="0" cellspacing="0" style="width: 100%">
                                        <tr align="left">
                                            <td>
                                                <cc1:TabContainer ID="TabContainer" runat="server" ActiveTabIndex="0">
                                                    <cc1:TabPanel ID="tpProductInfo" runat="server" HeaderText="Product">
                                                        <HeaderTemplate>
                                                            Product Info
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <br />
                                                            <asp:UpdatePanel ID="UpdatePanelProduct" runat="server" UpdateMode="Conditional">
                                                                <ContentTemplate>
                                                                    <asp:FormView ID="fvProductInfo" runat="server" DataKeyNames="ProductID" DataSourceID="dsProduct"
                                                                        Width="750px">
                                                                        <EditItemTemplate>
                                                                            <table width="100%">
                                                                                <tr>
                                                                                    <td>
                                                                                        ProductID:
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:Label ID="ProductIDLabel1" runat="server" Text='<%# Eval("ProductID") %>'></asp:Label></td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        Created By:
                                                                                    </td>
                                                                                    <td>
                                                                                        <%# eval("CreatedName") %>
                                                                                        on
                                                                                        <%# Eval("DateCreated") %>
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Brand:
                                                                                    </td>
                                                                                    <td>
                                                                                        <%# eval("BrandName") %>
                                                                                    </td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        Series:</td>
                                                                                    <td>
                                                                                    <asp:Label ID="lblProductCategoryID"  Visible ="false" Text ='<%# Eval("ProductCategoryID") %>' runat="server" ></asp:Label>
                                                                                        <asp:DropDownList ID="ddlSeries" runat="server" DataSourceID="dsSeries" DataTextField="Name"
                                                                                            DataValueField="ProductSeriesID" SelectedValue='<%# Bind("ProductSeriesID") %>'>
                                                                                        </asp:DropDownList>
                                                                                        <asp:SqlDataSource ID="dsSeries" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                                            SelectCommand="GetProductSeriesBySubCategoryID" SelectCommandType="StoredProcedure" CancelSelectOnNullParameter="true">
                                                                                            <SelectParameters >
                                                                                                <asp:ControlParameter Name ="SubCategoryID" Type ="int32" ControlID ="fvProductInfo$lblProductCategoryID" PropertyName ="Text" />
                                                                                            </SelectParameters>
                                                                                             
                                                                                        </asp:SqlDataSource>
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Category:</td>
                                                                                    <td colspan="4">
                                                                                        <%#Eval("EcommerceParentCategoryName")%>
                                                                                        &nbsp;>&nbsp;<%#Eval("EcommerceCategoryName")%></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        ProductModelNumber:
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="ProductModelNumberTextBox" runat="server" Text='<%# Bind("ProductModelNumber") %>'>
                                                                                        </asp:TextBox>
                                                                                    </td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        Year:</td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="YearTextBox" runat="server" Text='<%# Bind("Year") %>'>
                                                                                        </asp:TextBox></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Name:
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="NameTextBox" runat="server" Text='<%# Bind("Name") %>' Width="250px"></asp:TextBox>
                                                                                    </td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        Freight: &nbsp;
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:TextBox ID="FreightTextBox" runat="server" Text='<%# Bind("Freight") %>'>
                                                                                        </asp:TextBox></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Is Supported:</td>
                                                                                    <td>
                                                                                        <asp:CheckBox ID="IsSupportedCheckBox" runat="server" Checked='<%# Bind("IsSupported") %>' /></td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        Active
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:CheckBox ID="ActiveCheckBox" runat="server" Checked='<%# Bind("Active") %>' /></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Is Sellable:</td>
                                                                                    <td>
                                                                                        <asp:CheckBox Enabled="true" ID="IsSellableCheckBox" runat="server" Checked='<%# eval("IsSellable") %>' /></td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;
                                                                                    </td>
                                                                                    <td>
                                                                                        MSRP:</td>
                                                                                    <td>
                                                                                        <asp:Label ID="MSRPTextBox" runat="server" Text='<%# eval("MSRP") %>'>
                                                                                        </asp:Label></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Active Item on Mktg WebSite:</td>
                                                                                    <td>
                                                                                        <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# bind("IsActiveOnMktgSite") %>' /></td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;
                                                                                    </td>
                                                                                    <td>
                                                                                        Current Year Product:</td>
                                                                                    <td>
                                                                                        <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# bind("IsCurrentYearProduct") %>' /></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td style="width: 100%" colspan="5">
                                                                                        <br />
                                                                                        Description<br />
                                                                                        <br />
                                                                                        Use the HTML Description tab to edit this field
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td colspan="5" align="right">
                                                                                        <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="False" CommandName="Update"
                                                                                            Text="Update">
                                                                                        </asp:LinkButton>
                                                                                        <asp:LinkButton ID="UpdateCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                                                                                            Text="Cancel">
                                                                                        </asp:LinkButton></td>
                                                                                </tr>
                                                                            </table>
                                                                        </EditItemTemplate>
                                                                        <InsertItemTemplate>
                                                                            ProductID:
                                                                            <asp:TextBox ID="ProductIDTextBox" runat="server" Text='<%# Bind("ProductID") %>'>
                                                                            </asp:TextBox><br />
                                                                            ProductModelNumber:
                                                                            <asp:TextBox ID="ProductModelNumberTextBox" runat="server" Text='<%# Bind("ProductModelNumber") %>'>
                                                                            </asp:TextBox><br />
                                                                            Name:
                                                                            <asp:TextBox ID="NameTextBox" runat="server" Text='<%# Bind("Name") %>'>
                                                                            </asp:TextBox><br />
                                                                            Description:
                                                                            <asp:TextBox ID="DescriptionTextBox" runat="server" Text='<%# Bind("Description") %>'>
                                                                            </asp:TextBox><br />
                                                                            Year:
                                                                            <asp:TextBox ID="YearTextBox" runat="server" Text='<%# Bind("Year") %>'>
                                                                            </asp:TextBox><br />
                                                                            Freight:
                                                                            <asp:TextBox ID="FreightTextBox" runat="server" Text='<%# Bind("Freight") %>'>
                                                                            </asp:TextBox><br />
                                                                            MSRP:
                                                                            <asp:TextBox ID="MSRPTextBox" runat="server" Text='<%# Bind("MSRP") %>'>
                                                                            </asp:TextBox><br />
                                                                            Specifications:
                                                                            <asp:TextBox ID="SpecificationsTextBox" runat="server" Text='<%# Bind("Specifications") %>'>
                                                                            </asp:TextBox><br />
                                                                            PrimaryResourceImageID:
                                                                            <asp:TextBox ID="PrimaryResourceImageIDTextBox" runat="server" Text='<%# Bind("PrimaryResourceImageID") %>'>
                                                                            </asp:TextBox><br />
                                                                            PrimaryResourceImageID_large:
                                                                            <asp:TextBox ID="PrimaryResourceImageID_largeTextBox" runat="server" Text='<%# Bind("PrimaryResourceImageID_large") %>'>
                                                                            </asp:TextBox><br />
                                                                            PrimaryResourceImageID_small:
                                                                            <asp:TextBox ID="PrimaryResourceImageID_smallTextBox" runat="server" Text='<%# Bind("PrimaryResourceImageID_small") %>'>
                                                                            </asp:TextBox><br />
                                                                            ImagePath:
                                                                            <asp:TextBox ID="ImagePathTextBox" runat="server" Text='<%# Bind("ImagePath") %>'>
                                                                            </asp:TextBox><br />
                                                                            ImagePath_small:
                                                                            <asp:TextBox ID="ImagePath_smallTextBox" runat="server" Text='<%# Bind("ImagePath_small") %>'>
                                                                            </asp:TextBox><br />
                                                                            ImagePath_large:
                                                                            <asp:TextBox ID="ImagePath_largeTextBox" runat="server" Text='<%# Bind("ImagePath_large") %>'>
                                                                            </asp:TextBox><br />
                                                                            CreatedName:
                                                                            <asp:TextBox ID="CreatedNameTextBox" runat="server" Text='<%# Bind("CreatedName") %>'>
                                                                            </asp:TextBox><br />
                                                                            DateCreated:
                                                                            <asp:TextBox ID="DateCreatedTextBox" runat="server" Text='<%# Bind("DateCreated") %>'>
                                                                            </asp:TextBox><br />
                                                                            IsSupported:
                                                                            <asp:CheckBox ID="IsSupportedCheckBox" runat="server" Checked='<%# Bind("IsSupported") %>' /><br />
                                                                            CreatedBy:
                                                                            <asp:TextBox ID="CreatedByTextBox" runat="server" Text='<%# Bind("CreatedBy") %>'>
                                                                            </asp:TextBox><br />
                                                                            IsSellable:
                                                                            <asp:CheckBox ID="IsSellableCheckBox" runat="server" Checked='<%# Bind("IsSellable") %>' /><br />
                                                                            ProductSeriesID:
                                                                            <asp:TextBox ID="ProductSeriesIDTextBox" runat="server" Text='<%# Bind("ProductSeriesID") %>'>
                                                                            </asp:TextBox><br />
                                                                            Active:
                                                                            <asp:CheckBox ID="ActiveCheckBox" runat="server" Checked='<%# Bind("Active") %>' /><br />
                                                                            ParentCategoryID:
                                                                            <asp:TextBox ID="ParentCategoryIDTextBox" runat="server" Text='<%# Bind("ParentCategoryID") %>'>
                                                                            </asp:TextBox><br />
                                                                            ProductCategoryID:
                                                                            <asp:TextBox ID="ProductCategoryIDTextBox" runat="server" Text='<%# Bind("ProductCategoryID") %>'>
                                                                            </asp:TextBox><br />
                                                                            BrandID:
                                                                            <asp:TextBox ID="BrandIDTextBox" runat="server" Text='<%# Bind("BrandID") %>'>
                                                                            </asp:TextBox><br />
                                                                            BrandName:
                                                                            <asp:TextBox ID="BrandNameTextBox" runat="server" Text='<%# Bind("BrandName") %>'>
                                                                            </asp:TextBox><br />
                                                                            SeriesName:
                                                                            <asp:TextBox ID="SeriesNameTextBox" runat="server" Text='<%# Bind("SeriesName") %>'>
                                                                            </asp:TextBox><br />
                                                                            ParentCategoryName:
                                                                            <asp:TextBox ID="ParentCategoryNameTextBox" runat="server" Text='<%# Bind("ParentCategoryName") %>'>
                                                                            </asp:TextBox><br />
                                                                            CategoryName:
                                                                            <asp:TextBox ID="CategoryNameTextBox" runat="server" Text='<%# Bind("CategoryName") %>'>
                                                                            </asp:TextBox><br />
                                                                            <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" CommandName="Insert"
                                                                                Text="Insert">
                                                                            </asp:LinkButton>
                                                                            <asp:LinkButton ID="InsertCancelButton" runat="server" CausesValidation="False" CommandName="Cancel"
                                                                                Text="Cancel">
                                                                            </asp:LinkButton>
                                                                        </InsertItemTemplate>
                                                                        <ItemTemplate>
                                                                            <table width="100%">
                                                                                <tr>
                                                                                    <td>
                                                                                        ProductID:
                                                                                        
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:Label ID="ProductIDLabel1" runat="server" Text='<%# Eval("ProductID") %>'></asp:Label></td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        Created By:
                                                                                    </td>
                                                                                    <td>
                                                                                        <%# eval("CreatedName") %>
                                                                                        on
                                                                                        <%# Eval("DateCreated") %>
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Brand:
                                                                                    </td>
                                                                                    <td>
                                                                                        <%# eval("BrandName") %>
                                                                                    </td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        Series:</td>
                                                                                    <td>
                                                                                        <%# eval("SeriesName") %>
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Category:</td>
                                                                                    <td colspan="4">
                                                                                        <%#Eval("EcommerceParentCategoryName")%>
                                                                                        &nbsp;>&nbsp;<%#Eval("EcommerceCategoryName")%></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        ProductModelNumber:
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:Label ID="ProductModelNumberTextBox" runat="server" Text='<%# Bind("ProductModelNumber") %>'>
                                                                                        </asp:Label>
                                                                                    </td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        </td>
                                                                                    <td>
                                                                                        <asp:Label ID="YearTextBox" visible ="false" runat="server" Text='<%# Bind("Year") %>'>
                                                                                        </asp:Label></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>
                                                                                        Name:
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:Label ID="NameTextBox" runat="server" Text='<%# Bind("Name") %>' Width="250px"></asp:Label>
                                                                                    </td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;</td>
                                                                                    <td>
                                                                                        Freight: &nbsp;
                                                                                    </td>
                                                                                    <td>
                                                                                        <asp:Label ID="FreightTextBox" runat="server" Text='<%# Bind("Freight") %>'>
                                                                                        </asp:Label></td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>IsSupported:</td>
                                                                                    <td><asp:CheckBox Enabled="false" ID="IsSupportedCheckBox" runat="server" Checked='<%# Bind("IsSupported") %>' /></td>
                                                                                    <td style="width: 10px">&nbsp;</td>
                                                                                    <td>Active</td>
                                                                                    <td>
                                                                                        <asp:CheckBox ID="ActiveCheckBox" runat="server" Checked='<%# Bind("Active") %>' />
                                                                                        &nbsp;&nbsp;&nbsp;&nbsp;
                                                                                        <asp:LinkButton ID="UpdateActive" runat="server" Font-Size="XX-Small" OnClick="UpdateActive_OnClick" Text="Update"></asp:LinkButton>
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td>IsSellable:</td>
                                                                                    <td>
                                                                                        <asp:CheckBox ID="IsSellableCheckBox" Enabled="false"  runat="server" Checked='<%# eval("IsSellable") %>' />
                                                                                        &nbsp;&nbsp;&nbsp;&nbsp;
                                                                                        <asp:LinkButton ID="UpdateSellable" Visible="false" runat="server" Font-Size="XX-Small" OnClick="UpdateSellable_OnClick" Text="Update"></asp:LinkButton>
                                                                                    </td>
                                                                                    <td style="width: 10px">&nbsp;</td>
                                                                                    <td>
                                                                                        MSRP:</td>
                                                                                    <td>
                                                                                        <asp:Label ID="MSRPTextBox" runat="server" Text='<%# eval("MSRP") %>'>
                                                                                        </asp:Label></td>
                                                                                </tr>
                                                                               <%-- <tr>
                                                                                    <td>
                                                                                       <!-- Active Item on Mktg WebSite:--></td>
                                                                                    <td>
                                                                                        <asp:CheckBox Enabled="false" Visible ="false"  ID="CheckBox1" runat="server" Checked='<%# bind("IsActiveOnMktgSite") %>' /></td>
                                                                                    <td style="width: 10px">
                                                                                        &nbsp;
                                                                                    </td>
                                                                                    <td>
                                                                                        Current Year Product:</td>
                                                                                    <td>
                                                                                        <asp:CheckBox Enabled="false" ID="CheckBox2" runat="server" Checked='<%# bind("IsCurrentYearProduct") %>' /></td>
                                                                                </tr>--%>
                                                                                <tr>
                                                                                    <td style="width: 100%" colspan="5">
                                                                                        <br />
                                                                                        Description<br />
                                                                                        <br />
                                                                                        <asp:Label ID="DescriptionTextBox" runat="server" Text='<%# Bind("Description") %>'
                                                                                            Height="81px" Width="623px"></asp:Label>
                                                                                        &nbsp; &nbsp;</td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td colspan="5" align="right">
                                                                                        <asp:LinkButton ID="EditButton" Visible ="false"  runat="server" CausesValidation="True" CommandName="Edit"
                                                                                            Text="Edit">
                                                                                        </asp:LinkButton>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </ItemTemplate>
                                                                    </asp:FormView>
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                            <br />
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpProductDescription" runat="server" HeaderText="HTML Description">
                                                        <HeaderTemplate>
                                                            HTML Description
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <asp:UpdatePanel ID="upProductDescription" runat="server" UpdateMode="Conditional">
                                                                <ContentTemplate>
                                                                    <asp:FormView ID="fvHTMLDescription" runat="server" DataKeyNames="ProductID" DataSourceID="dsProductDescription"
                                                                        Width="750px" DefaultMode="edit">
                                                                        <EditItemTemplate>
                                                                            <table>
                                                                                <tr>
                                                                                    <td>
                                                                                        <FTB:FreeTextBox ID="FreeTextBox2" runat="Server" Text='<%# Bind("Description") %>' />
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td align="right">
                                                                                        <asp:LinkButton ID="UpdateButton" runat="server" CausesValidation="False" CommandName="Update"
                                                                                            Text="Update">
                                                                                        </asp:LinkButton>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </EditItemTemplate>
                                                                    </asp:FormView>
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpParts" runat="server" HeaderText="Parts">
                                                        <HeaderTemplate>
                                                            Parts
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <asp:UpdatePanel ID="UpdatePanelBOM" runat="server" UpdateMode="Conditional">
                                                                <ContentTemplate>
                                                                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="650">
                                                                        <tr>
                                                                            <td><asp:Panel ID="pnlUploadBOM" runat="server" GroupingText="Upload Product BOM">
                                                                                <table align="left" border="0" cellpadding="0" cellspacing="0" width="550">
                                                                                    
                                                                                    <tr>
                                                                                        <td class="FunctionData"><br />
                                                                                            All Existing Parts for this Model will be DELETED and reloaded from Uploaded BOM.
                                                                                        </td>
                                                                                    </tr>
                                                                                    <tr>
                                                                                        <td class="FunctionData"><br />
                                                                                            Select BOM File (Excel Spreadsheet ONLY)<br />
                                                                                            <input id="txtBOMFile" runat="server" size="45" type="file" />
                                                                                        </td>
                                                                                    </tr>
                                                                                    <tr>
                                                                                        <td align="right">
                                                                                            <asp:UpdatePanel ID="UpdatePanelParts" runat="server" UpdateMode="conditional">
                                                                                                <Triggers>
                                                                                                    <asp:PostBackTrigger ControlID="btnUploadBOM" />
                                                                                                </Triggers>
                                                                                                <ContentTemplate>
                                                                                                    <asp:Button ID="btnUploadBOM" runat="server" Text="Upload File" />
                                                                                                </ContentTemplate>
                                                                                            </asp:UpdatePanel>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table></asp:Panel>
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td align="left"><br /><br />
                                                                                <asp:Panel ID="pnlAddPart" runat="server" GroupingText="Add Existing Part to Product">
                                                                                    &nbsp;<asp:Label ID="lblPartError" CssClass="errmsg" runat="server"></asp:Label>
                                                                                    <asp:FormView ID="fvAddPart" runat="server" DataSourceID="dsAddProductPart" DataKeyNames="partid"
                                                                                        DefaultMode="Insert" >
                                                                                        <InsertItemTemplate >
                                                                                            <table>
                                                                                                
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        Part Number</td>
                                                                                                    <td>
                                                                                                        Warranty Duration</td>
                                                                                                    <td>
                                                                                                        Call Letter</td>
                                                                                                    <td>
                                                                                                        Qty</td>
                                                                                                    <td>
                                                                                                        &nbsp;</td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        <asp:TextBox ID="txtPartNumber" MaxLength="25" runat="server" >
                                                                                                        </asp:TextBox></td>
                                                                                                    <td>
                                                                                                        <asp:TextBox  ID="txtWarrantyDuration" MaxLength="5" runat="server" Text='<%# Bind("WarrantyDuration") %>'>
                                                                                                        </asp:TextBox></td>
                                                                                                    <td>
                                                                                                        <asp:TextBox ID="txtCallLetter" MaxLength="5" runat="server" Text='<%# Bind("CallLetter") %>'>
                                                                                                        </asp:TextBox></td>
                                                                                                    <td>
                                                                                                        <asp:TextBox ID="txtQuantity" MaxLength="2" runat="server" Text='<%# Bind("Quantity") %>'></asp:TextBox></td>
                                                                                                    <td>
                                                                                                        <asp:LinkButton ID="InsertButton" runat="server" CausesValidation="True" CommandName="Insert"
                                                                                                            Text="Insert">
                                                                                                        </asp:LinkButton></td>
                                                                                                </tr>
                                                                                            </table>
                                                                                            <br />
                                                                                        </InsertItemTemplate>
                                                                                    </asp:FormView>
                                                                                    <asp:SqlDataSource ID="dsPartSearch" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                        SelectCommand="[GetPartByPartNumber]" SelectCommandType="StoredProcedure">
                        <SelectParameters>
                        <asp:Parameter Name="PartNumber" DefaultValue="" Type="string"  />
                        </SelectParameters>
                        </asp:SqlDataSource>
                                                                                </asp:Panel>
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td>
                                                                                <br />
                                                                                <br />
                                                                                <asp:GridView ID="gvParts" runat="server" AutoGenerateColumns="False" DataKeyNames="PartProductID"
                                                                                    DataSourceID="dsPartLists" Width="100%" AllowSorting="True" >
                                                                                    <Columns>
                                                                                        <asp:BoundField DataField="CallLetter" HeaderText="Call Tag" SortExpression="CallLetter" />
                                                                                        <asp:TemplateField>
                                                                                            <ItemTemplate>
                                                                                                <asp:HyperLink ID="hlManagePart" runat="server" Text='<%# eval("PartNumber") %>'
                                                                                                    ToolTip="Manage this Part" NavigateUrl='<%# Eval("PartID", "ManagePart.aspx?PartID={0}") %>'></asp:HyperLink>
                                                                                            </ItemTemplate>
                                                                                        </asp:TemplateField>
                                                                                        <asp:BoundField DataField="Name" HeaderText="Name" ReadOnly="True" SortExpression="Name" />
                                                                                        <asp:TemplateField HeaderText="Price" SortExpression="MSRP">
                                                                                            <ItemTemplate>
                                                                                                <asp:Label ID="lblMSRP" runat="server" Text='<%# Eval("MSRP") %>'></asp:Label>
                                                                                            </ItemTemplate>
                                                                                        </asp:TemplateField>
                                                                                        <asp:TemplateField HeaderText="Warranty" SortExpression="WarrantyDuration">
                                                                                            <ItemTemplate>
                                                                                                <asp:Label ID="lblWarranty" runat="server" Text='<%# Eval("Duration") %>'></asp:Label>
                                                                                            </ItemTemplate>
                                                                                            <EditItemTemplate>
                                                                                                <asp:textBox ID="txtWarrantyDuration" runat="server" Text='<%# Bind("WarrantyDuration") %>'></asp:textBox>
                                                                                            </EditItemTemplate>
                                                                                        </asp:TemplateField>
                                                                                        
                                                                                        
                                                                                        <asp:BoundField DataField="Quantity" HeaderText="Quantity" ReadOnly="False" SortExpression="Quantity" InsertVisible="False" />
                                                                                        <asp:TemplateField ShowHeader="False">
                                                                                            <ItemTemplate>
                                                                                            
                                                                                                <asp:LinkButton ID="lnkDeleteProductPart" runat="server" CausesValidation="False"
                                                                                                    CommandName="Delete" OnClientClick="return  confirm('Are you sure you want to DELETE Part from this Product?');"
                                                                                                    Text="Delete"></asp:LinkButton>&nbsp;
                                                                                            </ItemTemplate>
                                                                                        </asp:TemplateField>
                                                                                        
                                                                                        <asp:BoundField DataField="PartProductID" HeaderText="PartProductID" InsertVisible="False"
                                                                                            ReadOnly="True" SortExpression="PartProductID" Visible="False" />
                                                                                        <asp:CommandField ShowEditButton="True" ControlStyle-ForeColor="#2c71b6" ControlStyle-Font-Bold="true" ControlStyle-Font-Names="'trebuchet ms',trebuchet,arial,verdana,sans-serif" ControlStyle-Font-Size="12px" />
                                                                                            
                                                                                    </Columns>
                                                                                </asp:GridView>
                                                                                <asp:SqlDataSource ID="dsAddProductPart" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                                    InsertCommand="AddPartProductJUNC" InsertCommandType="StoredProcedure" SelectCommand="GetPart"
                                                                                    SelectCommandType="StoredProcedure">
                                                                                    
                                                                                    <InsertParameters>
                                                                                        <asp:Parameter Name="PartID" Type="Int32" />
                                                                                        <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="Int32" />
                                                                                        <asp:Parameter Name="WarrantyDuration" Type="Int32" DefaultValue="12" />
                                                                                        <asp:Parameter Name="CallLetter" Type="String" />
                                                                                        <asp:Parameter Name="Quantity" Type="Int32" DefaultValue="1"  />
                                                                                    </InsertParameters>
                                                                                </asp:SqlDataSource>
                                                                                <asp:SqlDataSource ID="dsPartLists" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                                    SelectCommand="GetProductParts" SelectCommandType="StoredProcedure" CancelSelectOnNullParameter="true"
                                                                                    DeleteCommand="DeletePartFromProduct" DeleteCommandType="storedprocedure" UpdateCommand="UpdateProductPart" UpdateCommandType="storedProcedure" >
                                                                                    <SelectParameters>
                                                                                        <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="Int32" />
                                                                                    </SelectParameters>
                                                                                    <UpdateParameters>
                                                                                        <asp:ControlParameter ControlID="gvParts" Name="PartProductID" PropertyName="SelectedValue" Type="Int32" />
                                                                                        <asp:Parameter Name="CallLetter" Size="5" Type="String" />
                                                                                        <asp:Parameter Name="WarrantyDuration" Type="Int32" DefaultValue="12" />
                                                                                        <asp:Parameter Name="Quantity" Type="Int32" DefaultValue="1"  />
                                                                                    </UpdateParameters>
                                                                                    <DeleteParameters>
                                                                                        <asp:ControlParameter ControlID="gvParts" Name="PartProductID" PropertyName="SelectedValue"
                                                                                            Type="Int32" />
                                                                                    </DeleteParameters>
                                                                                </asp:SqlDataSource>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpProductDiagrams" runat="server" HeaderText="Diagrams">
                                                        <HeaderTemplate>
                                                            Product Diagrams
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <asp:UpdatePanel ID="UpdatePanelDiagramMain" runat="server" Visible="true" UpdateMode="conditional">
                                                                <ContentTemplate>
                                                                    <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender10" runat="server"
                                                                        TargetControlID="TabContainer$tpProductDiagrams$UpdatePanelDiagram">
                                                                    </cc1:UpdatePanelAnimationExtender>
                                                                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="650">
                                                                        <tr>
                                                                            <td class="FunctionTitles">
                                                                                <br />
                                                                                Upload Parts Diagram &nbsp;&nbsp;
                                                                                <br />
                                                                                <br />
                                                                                <asp:Label ID="lblCurrentDiagram" runat="server"></asp:Label>
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td class="FunctionData">
                                                                                <br />
                                                                                <br />
                                                                                Any Existing Parts Diagram for this Model will be DELETED and replaced with Uploaded
                                                                                Diagram.
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td>
                                                                                <table align="center" border="0" cellpadding="0" cellspacing="0" width="450">
                                                                                    <tr>
                                                                                        <td class="FunctionData">
                                                                                            <br />
                                                                                            <br />
                                                                                            Select Parts Diagram (Images ONLY)<br />
                                                                                            <br />
                                                                                            <input id="txtDiagramFile" runat="server" size="45" type="file" />
                                                                                        </td>
                                                                                    </tr>
                                                                                    <tr>
                                                                                        <td align="right">
                                                                                            <asp:UpdatePanel ID="UpdatePanelDiagram" runat="server" UpdateMode="conditional">
                                                                                                <Triggers>
                                                                                                    <asp:PostBackTrigger ControlID="btnUploadDiagram" />
                                                                                                </Triggers>
                                                                                                <ContentTemplate>
                                                                                                    <asp:Button ID="btnUploadDiagram" runat="server" Text="Upload File" /></ContentTemplate>
                                                                                            </asp:UpdatePanel>
                                                                                        </td>
                                                                                    </tr>
                                                                                    <tr>
                                                                                        <td align="right">
                                                                                            <br />
                                                                                            <asp:FormView ID="fvDiagram" runat="server" DataSourceID="dsDiagram">
                                                                                                <ItemTemplate>
                                                                                                    <asp:ImageButton ID="btnDiagram" ImageUrl='<%# Utilities.GetImageURL(Eval("filepath") & Eval("filename")) %>'
                                                                                                        runat="server" Height="500px" Visible="False" Width="550px" /><br />
                                                                                                </ItemTemplate>
                                                                                            </asp:FormView>
                                                                                            <asp:SqlDataSource ID="dsDiagram" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                                                SelectCommand="[GetProductDiagram]" SelectCommandType="StoredProcedure">
                                                                                                <SelectParameters>
                                                                                                    <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="int32" />
                                                                                                </SelectParameters>
                                                                                            </asp:SqlDataSource>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpProductImages" runat="server" HeaderText="Images">
                                                        <HeaderTemplate>
                                                            Images
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <asp:UpdatePanel ID="UpdatePanelImages" runat="server" Visible="true" UpdateMode="conditional">
                                                                <ContentTemplate>
                                                                    &nbsp;<asp:LinkButton ID="lnkPrimary" runat="server" Text="Primary Product Image"></asp:LinkButton>
                                                                    &nbsp; | &nbsp;
                                                                    <asp:LinkButton ID="lnkAlternate" runat="server" Text="Alternate Product Image Views"></asp:LinkButton>
                                                                    <br />
                                                                    <br />
                                                                    <asp:Panel ID="pPrimary" runat="server" Visible="false">
                                                                        <div style="border-right: black 1px solid; border-top: black 1px solid; border-left: black 1px solid;
                                                                            border-bottom: black 1px solid">
                                                                            <table border="0">
                                                                                <tr>
                                                                                    <td>
                                                                                        &nbsp;
                                                                                        <asp:Image ID="imgPrimary" runat="server" /></td>
                                                                                    <td align="center" valign="top">
                                                                                        <asp:CheckBox ID="chkAutoResize" runat="server" AutoPostBack="True" Checked="true"
                                                                                            Text="Enable auto resize" ToolTip="This feature allows for automatic creation of thumbnail and regular image size based on image properites. " />
                                                                                        <asp:Panel ID="pImageSizeSettings" runat="server" Visible="true">
                                                                                            <table bgcolor="#eeeeee" border="0" width="350">
                                                                                                <tr>
                                                                                                    <td align="center" colspan="3">
                                                                                                        <b>Image Resize Properites</b></td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        &nbsp;</td>
                                                                                                    <td>
                                                                                                        Width</td>
                                                                                                    <td>
                                                                                                        Height</td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        Feature:</td>
                                                                                                    <td>
                                                                                                        <asp:Label ID="lblFeatureImageSizeWidth" runat="server"></asp:Label>
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <asp:Label ID="lblFeatureImageSizeHeight" runat="server"></asp:Label>
                                                                                                    </td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        Small:</td>
                                                                                                    <td>
                                                                                                        <asp:Label ID="lblSmallImageSizeWidth" runat="server"></asp:Label>
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <asp:Label ID="lblSmallImageSizeHeight" runat="server"></asp:Label>
                                                                                                    </td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        Regular:
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <asp:Label ID="lblRegularImageSizeWidth" runat="server"></asp:Label>
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <asp:Label ID="lblRegularImageSizeHeight" runat="server"></asp:Label></td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        Large:
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <asp:Label ID="lblLargeImageSizeWidth" runat="server"></asp:Label>
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <asp:Label ID="lblLargeImageSizeHeight" runat="server"></asp:Label></td>
                                                                                                </tr>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        Constrain Image on Resize:
                                                                                                    </td>
                                                                                                    <td colspan ="2" align="left" >
                                                                                                        <asp:Label ID="lblConstrain" runat="server" Text =""></asp:Label>
                                                                                                    </td>
                                                                                                   
                                                                                                </tr>
                                                                                                
                                                                                                <tr>
                                                                                                    <td align="center" colspan="3">
                                                                                                        <br />
                                                                                                        <asp:FileUpload ID="FileProductImageUpload" runat="server" /><br />
                                                                                                        <br />
                                                                                                        <asp:UpdatePanel ID="UpdatePanel1" runat="server" UpdateMode="conditional">
                                                                                                            <Triggers>
                                                                                                                <asp:PostBackTrigger ControlID="btnUpdateImage" />
                                                                                                                <asp:PostBackTrigger ControlID="btnRegenerateImage" />
                                                                                                            </Triggers>
                                                                                                            <ContentTemplate>
                                                                                                                <asp:Button ID="btnUpdateImage" runat="server" OnClientClick="return  confirm('This will replace all existing product image sizes. Are you sure');"
                                                                                                                    Text="Update Image" />
                                                                                                                <asp:Button ID="btnRegenerateImage" runat="server" OnClientClick="return  confirm('This will replace all existing product image sizes. Are you sure');"
                                                                                                                    Text="Reload Default Images" ToolTip="This option generate resized image based on the current largest image specified." />
                                                                                                                <br />
                                                                                                                <br />
                                                                                                            </ContentTemplate>
                                                                                                        </asp:UpdatePanel>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </asp:Panel>
                                                                                    </td>
                                                                                </tr>
                                                                                <tr>
                                                                                    <td align="center" colspan="3">
                                                                                        <!-- 3 Image Table !-->
                                                                                        <asp:Panel ID="pProductImages" runat="server" Visible="false">
                                                                                            <table>
                                                                                                <tr>
                                                                                                    <td>
                                                                                                        <table>
                                                                                                            <tr>
                                                                                                                <td align="center">
                                                                                                                    Featured Image</td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td align="center">
                                                                                                                    <asp:Image ID="imgFeature" runat="server" />
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td>
                                                                                                                    <asp:FileUpload ID="FileFeatureImageUpload" runat="server" />
                                                                                                                    <br />
                                                                                                                    <asp:UpdatePanel ID="UpdatePanel3" runat="server" UpdateMode="conditional">
                                                                                                                        <Triggers>
                                                                                                                            <asp:PostBackTrigger ControlID="btnUpdateFeatureImage" />
                                                                                                                        </Triggers>
                                                                                                                        <ContentTemplate>
                                                                                                                            <asp:Button ID="btnUpdateFeatureImage" runat="server" Text="Update Image" />
                                                                                                                        </ContentTemplate>
                                                                                                                    </asp:UpdatePanel>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                        <br />
                                                                                                        <table>
                                                                                                            <tr>
                                                                                                                <td align="center">
                                                                                                                    Small Image</td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td align="center">
                                                                                                                    <asp:Image ID="imgSmall" runat="server" />
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td>
                                                                                                                    <asp:FileUpload ID="FileSmallImageUpload" runat="server" />
                                                                                                                    <br />
                                                                                                                    <asp:UpdatePanel ID="UpdatePanel4" runat="server" UpdateMode="conditional">
                                                                                                                        <Triggers>
                                                                                                                            <asp:PostBackTrigger ControlID="btnUpdateSmallImage" />
                                                                                                                        </Triggers>
                                                                                                                        <ContentTemplate>
                                                                                                                            <asp:Button ID="btnUpdateSmallImage" runat="server" Text="Update Image" />
                                                                                                                        </ContentTemplate>
                                                                                                                    </asp:UpdatePanel>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <table>
                                                                                                            <tr>
                                                                                                                <td align="center">
                                                                                                                    Regular Image
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td align="center">
                                                                                                                    <asp:Image ID="imgRegular" runat="server" />
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td>
                                                                                                                    <asp:FileUpload ID="FileRegularImageUpload" runat="server" />
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td>
                                                                                                                    <asp:UpdatePanel ID="UpdatePanel6" runat="server" UpdateMode="conditional">
                                                                                                                        <Triggers>
                                                                                                                            <asp:PostBackTrigger ControlID="btnUpdateRegularImage" />
                                                                                                                        </Triggers>
                                                                                                                        <ContentTemplate>
                                                                                                                            <asp:Button ID="btnUpdateRegularImage" runat="server" Text="Update Image" />
                                                                                                                        </ContentTemplate>
                                                                                                                    </asp:UpdatePanel>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                    <td>
                                                                                                        <table>
                                                                                                            <tr>
                                                                                                                <td align="center">
                                                                                                                    Large Image</td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td align="center">
                                                                                                                    <asp:Image ID="imgLarge" runat="server" />
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                            <tr>
                                                                                                                <td>
                                                                                                                    <asp:FileUpload ID="FileLargeImageUpload" runat="server" />
                                                                                                                    <br />
                                                                                                                    <asp:UpdatePanel ID="UpdatePanel7" runat="server" UpdateMode="conditional">
                                                                                                                        <Triggers>
                                                                                                                            <asp:PostBackTrigger ControlID="btnUpdateLargeImage" />
                                                                                                                        </Triggers>
                                                                                                                        <ContentTemplate>
                                                                                                                            <asp:Button ID="btnUpdateLargeImage" runat="server" Text="Update Image" />
                                                                                                                        </ContentTemplate>
                                                                                                                    </asp:UpdatePanel>
                                                                                                                </td>
                                                                                                            </tr>
                                                                                                        </table>
                                                                                                    </td>
                                                                                                </tr>
                                                                                            </table>
                                                                                        </asp:Panel>
                                                                                    </td>
                                                                                </tr>
                                                                            </table>
                                                                        </div>
                                                                    </asp:Panel>
                                                                    <asp:Panel ID="pAlternate" runat="server" Visible="false">
                                                                        Alternate Image Views<br />
                                                                        
                                                                        <asp:Panel ID="Panel1" runat="server" Visible ="true" >
                        <table border="0" bgcolor="#eeeeee" width ="350">
                            <tr><td colspan ="3" align="center" ><b>Image Resize Properites</b></td></tr>
                            <tr><td>&nbsp;</td><td>Width</td><td>Height</td></tr>
                            <tr>
                                <td>
                                    Alternate Regular:</td>
                                <td >
                                    
                                    <asp:label ID="lblAlternateRegularSizeWidth" runat="server" ></asp:label>
                                    
                                   </td><td>
                                    <asp:label ID="lblAlternateRegularSizeHeight" runat="server" ></asp:label>
                                    </td> 
                                    
                            </tr>
                            <tr>
                                <td>
                                    Alternate Small:</td>
                                <td >
                                    
                                    <asp:label ID="lblAlternateSmallSizeWidth" runat="server" ></asp:label>
                                    
</td><td>
                                    <asp:label ID="lblAlternateSmallSizeHeight" runat="server" ></asp:label>
         </td>                            
                            </tr>
                            
                            <tr>
                        <td colspan="3" align="center"><br />
                         <asp:FileUpload ID="FileUploadAlternate" runat="server" />
                                                                        <asp:UpdatePanel ID="UpdatePanel5" runat="server" UpdateMode="conditional">
                                                                            <Triggers>
                                                                                <asp:PostBackTrigger ControlID="btnAddAlternateImage" />
                                                                            </Triggers>
                                                                            <ContentTemplate>
                                                                                <asp:Button ID="btnAddAlternateImage" runat="server" Text="Add Image" />
                                                                            </ContentTemplate>
                                                                        </asp:UpdatePanel>
                            </td>
            </tr></table> 
                    </asp:Panel>
                    
                    
                                                                        
                                                                        
                                                                        
                                                                        <asp:GridView ID="gvAlternateImages" runat="server" AllowPaging="True" AutoGenerateColumns="False"
                                                                            AutoGenerateEditButton="True" DataKeyNames="ImageID" DataSourceID="dsAlternateImages">
                                                                            <Columns>
                                                                                <asp:TemplateField ShowHeader="False">
                                                                                    <EditItemTemplate>
                                                                                        <asp:Label ID="lblEditImageID" runat="server" Text='<%# Eval("ImageID") %>'></asp:Label>
                                                                                    </EditItemTemplate>
                                                                                    <ItemTemplate>
                                                                                        <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" CommandArgument='<%# Eval("ImageID") %>'
                                                                                            CommandName="Delete" OnClientClick="return confirm('You are about to delete this image. Are you sure?');"
                                                                                            Text="Delete"></asp:LinkButton>
                                                                                    </ItemTemplate>
                                                                                </asp:TemplateField>
                                                                                <asp:BoundField DataField="ImageID" HeaderText="ImageID" ReadOnly="True" SortExpression="ImageID"
                                                                                    Visible="False" />
                                                                                <asp:TemplateField HeaderText="Name" SortExpression="Name">
                   
                    <ItemTemplate>
                        <asp:Image ID="Image2" AlternateText ='<%# Eval("Description") %>' ImageAlign ="Middle" runat="server" ImageUrl = '<%# Utilities.GetAlternateThumbnailImagePath (Eval("ImagePath")) %>' />                        
                    </ItemTemplate>
                </asp:TemplateField>
                                                                                <asp:TemplateField HeaderText="Description" SortExpression="Description">
                                                                                    <EditItemTemplate>
                                                                                        <asp:TextBox ID="txtEditDescription" runat="server" Text='<%# Bind("Description") %>'></asp:TextBox>
                                                                                    </EditItemTemplate>
                                                                                    <ItemTemplate>
                                                                                        <asp:Label ID="Label2" runat="server" Text='<%# Bind("Description") %>'></asp:Label>
                                                                                    </ItemTemplate>
                                                                                </asp:TemplateField>
                                                                                <asp:TemplateField HeaderText="ImagePath" SortExpression="ImagePath" Visible="False">
                                                                                    <ItemTemplate>
                                                                                        <asp:Label ID="lblImagePath" runat="server" Text='<%# Bind("ImagePath") %>'></asp:Label>
                                                                                    </ItemTemplate>
                                                                                </asp:TemplateField>
                                                                                <asp:TemplateField HeaderText="DateCreated" SortExpression="DateCreated">
                                                                                    <EditItemTemplate>
                                                                                        <asp:Label ID="Label5" runat="server" Text='<%# Bind("DateCreated") %>'></asp:Label>
                                                                                    </EditItemTemplate>
                                                                                    <ItemTemplate>
                                                                                        <asp:Label ID="Label4" runat="server" Text='<%# Bind("DateCreated") %>'></asp:Label>
                                                                                    </ItemTemplate>
                                                                                </asp:TemplateField>
                                                                            </Columns>
                                                                        </asp:GridView>
                                                                        <asp:SqlDataSource ID="dsAlternateImages" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                            DeleteCommand="DeleteImage" DeleteCommandType="StoredProcedure" SelectCommand="SELECT Image.ImageID, Image.Description, Image.ImagePath, Image.DateCreated FROM Image INNER JOIN ImageResourceJUNC ON Image.ImageID = ImageResourceJUNC.ImageID WHERE (ImageResourceJUNC.ResourceTypeID = 4) AND (ImageResourceJUNC.ResourceID = @ProductID)"
                                                                            UpdateCommand="Update [Image] Set Description = @Description Where ImageID = @ImageID"
                                                                            >
                                                                            <SelectParameters>
                                                                                <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID"  />
                                                                            </SelectParameters>
                                                                            <DeleteParameters>
                                                                                <asp:ControlParameter ControlID="gvAlternateImages" Name="ImageID" PropertyName="SelectedValue"
                                                                                    Type="Int32" />
                                                                            </DeleteParameters>
                                                                            <UpdateParameters>
                                                                                <asp:Parameter Name="Description" Type="String" />
                                                                                <asp:Parameter Name="ImageID" Type="Int32" />
                                                                            </UpdateParameters>
                                                                        </asp:SqlDataSource>
                                                                    </asp:Panel>
                                                                    <div id="divImagePreview" style="display: none">
                                                                        <img id="imgPreview" alt="" src="" />
                                                                    </div>
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpProductSpecs" runat="server" HeaderText="Specs">
                                                        <HeaderTemplate>
                                                            Product Specs
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <asp:UpdatePanel ID="UpdatePanelSpecs" runat="server" UpdateMode="Conditional">
                                                                <ContentTemplate>
                                                                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="450">
                                                                        <tr>
                                                                            <td class="FunctionData">
                                                                                All Existing SPECS for this Model will be DELETED and reloaded from the uploaded
                                                                                spreadsheet.
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td class="FunctionData">
                                                                                Select SPECS File (Excel Spreadsheet ONLY)<br />
                                                                                <input id="txtSPECSFile" runat="server" size="45" type="file" />
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td align="right">
                                                                                <asp:UpdatePanel ID="UpdatePanelSPECUpload" runat="server" UpdateMode="conditional">
                                                                                    <Triggers>
                                                                                        <asp:PostBackTrigger ControlID="btnUploadSPECS" />
                                                                                    </Triggers>
                                                                                    <ContentTemplate>
                                                                                        <asp:Button ID="btnUploadSPECS" runat="server" Text="Upload File" OnClick="btnUploadSPECS_Click" />
                                                                                    </ContentTemplate>
                                                                                </asp:UpdatePanel>
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td align="right">
                                                                                <br />
                                                                                <br />
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                    <asp:GridView Width="500" ID="gvProductSpecs" DataKeyNames="ProductSpecificationID"
                                                                        runat="server" DataSourceID="dsSpecs">
                                                                        <Columns>
                                                                            <asp:CommandField ShowEditButton="True" />
                                                                            <asp:TemplateField ShowHeader="False">
                                                                                <ItemTemplate>
                                                                                    <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" CommandName="Delete"
                                                                                        Text="Clear Value" OnClientClick="confirm('You are about to clear this value. Are you sure?');"></asp:LinkButton>
                                                                                </ItemTemplate>
                                                                            </asp:TemplateField>
                                                                            <asp:BoundField DataField="ProductID" ReadOnly="True" Visible="False" />
                                                                            <asp:BoundField DataField="ProductSpecificationID" ReadOnly="True" Visible="False" />
                                                                            <asp:BoundField DataField="Name" ReadOnly="True" HeaderText="NAME" SortExpression="Name" />
                                                                            <asp:BoundField DataField="ProductSpecificationValue" HeaderText="VALUE" SortExpression="ProductSpecificationValue" />
                                                                        </Columns>
                                                                    </asp:GridView>
                                                                    <asp:SqlDataSource ID="dsSpecs" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                        SelectCommand="GetAllProductSpecifications" SelectCommandType="StoredProcedure"
                                                                        UpdateCommand="UpdateProductSpecification" UpdateCommandType="StoredProcedure"
                                                                        DeleteCommand="DeleteProductSpecification" DeleteCommandType="StoredProcedure">
                                                                        <DeleteParameters>
                                                                            <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" />
                                                                            <asp:Parameter Name="ProductSpecificationID" Type="Int32" />
                                                                        </DeleteParameters>
                                                                        <UpdateParameters>
                                                                            <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" />
                                                                            <asp:Parameter Name="ProductSpecificationID" Type="Int32" />
                                                                            <asp:Parameter Name="ProductSpecificationValue" Type="String" />
                                                                        </UpdateParameters>
                                                                        <SelectParameters>
                                                                            <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="Int32" />
                                                                        </SelectParameters>
                                                                    </asp:SqlDataSource>
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpProductManuals" runat="server" HeaderText="Manuals">
                                                        <HeaderTemplate>
                                                            Product Manuals
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <asp:UpdatePanel ID="UpdatePanelProductManuals" runat="server" Visible="true" UpdateMode="conditional">
                                                                <ContentTemplate>
                                                                    <br />
                                                                    <table border="0" cellpadding="0" cellspacing="0">
                                                                        <tr>
                                                                            <td style="width: 163px">
                                                                                Add New Product Manual
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width: 163px">
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width: 163px; height: 19px">
                                                                                Title (maximum length 250)</td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width: 163px">
                                                                                <asp:TextBox ID="txtProductManualTitle" runat="server" MaxLength="250" Width="365px"></asp:TextBox>
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width: 163px">
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width: 163px">
                                                                                Select Product Manual (PDF Only)</td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width: 163px">
                                                                                <input id="txtManualFile" runat="server" size="45" type="file" />
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width: 163px">
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td style="width: 163px">
                                                                                <asp:UpdatePanel ID="UpdatePanelUploadManual" runat="server" UpdateMode="conditional">
                                                                                    <Triggers>
                                                                                        <asp:PostBackTrigger ControlID="btnUploadManual" />
                                                                                    </Triggers>
                                                                                    <ContentTemplate>
                                                                                        <asp:Button ID="btnUploadManual" runat="server" Text="Upload File" /></ContentTemplate>
                                                                                </asp:UpdatePanel>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                    <br />
                                                                    <br />
                                                                    <asp:GridView ID="gvManuals" runat="server" AutoGenerateColumns="False" CssClass="resultsContent"
                                                                        DataKeyNames="KnowledgeID" DataSourceID="dsProductGuides" GridLines="None" Width="50%">
                                                                        <HeaderStyle CssClass="resultsHeader" Height="30px" />
                                                                        <EmptyDataTemplate>
                                                                            There are currently no Product Guides listed for this model
                                                                        </EmptyDataTemplate>
                                                                        <Columns>
                                                                            <asp:BoundField DataField="KnowledgeID" HeaderText="KnowledgeID" InsertVisible="False"
                                                                                ReadOnly="True" SortExpression="KnowledgeID" Visible="False" />
                                                                            <asp:TemplateField>
                                                                                <HeaderTemplate>
                                                                                    Product Guide(s)
                                                                                </HeaderTemplate>
                                                                                <ItemTemplate>
                                                                                    <asp:Image ID="Image1" runat="server" ImageUrl="~/Images/pdf.gif" /><asp:HyperLink
                                                                                        ID="hlManual" runat="server" CssClass="standardLink" NavigateUrl='<%# Eval("KnowledgeID", "http://www.tackleservice.com/consumer/ViewKnowledge.aspx?KnowledgeID={0}") %>'
                                                                                        Target="blank" Text='<%# trim(Eval("Title")) %>'></asp:HyperLink>
                                                                                </ItemTemplate>
                                                                            </asp:TemplateField>
                                                                            <asp:TemplateField ShowHeader="False">
                                                                                <ItemTemplate>
                                                                                    <asp:LinkButton ID="lnkDeleteManual" runat="server" CausesValidation="False" CommandName="Delete"
                                                                                        Text="Delete"></asp:LinkButton>
                                                                                    <cc1:ConfirmButtonExtender ID="ConfirmButtonExtender1" runat="server" ConfirmOnFormSubmit="True"
                                                                                        ConfirmText="Are You Sure You want to delete this MANUAL?" TargetControlID="lnkDeleteManual">
                                                                                    </cc1:ConfirmButtonExtender>
                                                                                </ItemTemplate>
                                                                            </asp:TemplateField>
                                                                        </Columns>
                                                                    </asp:GridView>
                                                                    <asp:SqlDataSource ID="dsProductGuides" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                        SelectCommand="[GetKnowledgeByResourceID]" SelectCommandType="StoredProcedure"
                                                                        DeleteCommand="DeleteProductManual" DeleteCommandType="StoredProcedure" >
                                                                        <DeleteParameters>
                                                                         <asp:ControlParameter ControlID="gvmanuals" Name="KnowledgeID" PropertyName="SelectedValue" Type="Int32" />
                                                                           
                                                                        </DeleteParameters>
                                                                        <SelectParameters>
                                                                            <asp:QueryStringParameter Name="ResourceID" QueryStringField="ProductID" Type="Int32" />
                                                                            <asp:Parameter DefaultValue="4" Name="ResourceTypeID" />
                                                                            <asp:Parameter DefaultValue="1" Name="KnowledgeTypeID" />
                                                                            <asp:Parameter DefaultValue="1" Name="KnowledgeTypeDirection" />
                                                                        </SelectParameters>
                                                                    </asp:SqlDataSource>
                                                                    <br />
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpProductRetailers" runat="server" HeaderText="UniversalParts">
                                                        <HeaderTemplate>
                                                            Product Retailers
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <asp:UpdatePanel ID="UpdatePanelRetailers" runat="server" UpdateMode="Conditional">
                                                                <ContentTemplate>
                                                                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="650">
                                                                        <tr>
                                                                            <td class="FunctionTitles">
                                                                                <br />
                                                                                Manage Product Retailers&nbsp;&nbsp;
                                                                                <asp:Label ID="Label6" runat="server"></asp:Label></td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td class="FunctionData">
                                                                                Add or Remove Retailers for this Product.<br />
                                                                                (Item will automatically show Lamplight as Retailer IF APPLICABLE)
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td>
                                                                                <table>
                                                                                    <tr>
                                                                                        <td style="width: 182px">
                                                                                            <br />
                                                                                            All Retailers<br />
                                                                                            <asp:ListBox ID="lbAvailableRetailers" runat="server" DataSourceID="dsAllRetailers"
                                                                                                DataTextField="Name" DataValueField="RetailerID" SelectionMode="Multiple" Height="139px"
                                                                                                Width="159px"></asp:ListBox>
                                                                                        </td>
                                                                                        <td style="width: 3px">
                                                                                            <asp:UpdatePanel ID="UpdatePanelRetailerButtons" runat="server" UpdateMode="conditional">
                                                                                                <ContentTemplate>
                                                                                                    <asp:Button ID="btnRetailerMoveRight" runat="server" Text=">>>>" />
                                                                                                    <asp:Button ID="btnRetailerMoveLeft" runat="server" Text="<<<<" /></ContentTemplate>
                                                                                            </asp:UpdatePanel>
                                                                                        </td>
                                                                                        <td>
                                                                                            Retailers for THIS Product<br />
                                                                                            <asp:ListBox ID="lbProductRetailers" runat="server" DataSourceID="dsProductRetailers"
                                                                                                DataTextField="Name" DataValueField="RetailerID" SelectionMode="Multiple" Height="139px"
                                                                                                Width="159px"></asp:ListBox>
                                                                                            <asp:SqlDataSource ID="dsProductRetailers" runat="server" CacheDuration="Infinite"
                                                                                                EnableCaching="true" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                                                SelectCommand="[GetProductRetailers]" SelectCommandType="StoredProcedure" UpdateCommand="[MANAGERUpdateProductRetailer]"
                                                                                                UpdateCommandType="storedProcedure">
                                                                                                <UpdateParameters>
                                                                                                    <asp:Parameter Name="ProductID" Type="Int32" />
                                                                                                    <asp:Parameter Name="RetailerID" Type="Int32" />
                                                                                                    <asp:Parameter Name="IsProductRetailer" Type="boolean" />
                                                                                                </UpdateParameters>
                                                                                                <SelectParameters>
                                                                                                    <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="int32" />
                                                                                                </SelectParameters>
                                                                                            </asp:SqlDataSource>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpCrossells" runat="server" HeaderText="Crossells">
                                                        <HeaderTemplate>
                                                            Cross Sells
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                            <asp:UpdatePanel ID="UpdatePanelCrossells" runat="server" UpdateMode="Conditional">
                                                                <ContentTemplate>
                                                                    <table align="center" border="0" cellpadding="0" cellspacing="0" width="650">
                                                                        <tr>
                                                                            <td class="FunctionTitles">
                                                                                <br />
                                                                                Manage Cross Sells&nbsp;&nbsp;
                                                                                <asp:Label ID="Label8" runat="server"></asp:Label></td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td class="FunctionData">
                                                                                Add or Remove Cross Sells for this Product.
                                                                            </td>
                                                                        </tr>
                                                                        <tr>
                                                                            <td>
                                                                                <table>
                                                                                    <tr>
                                                                                        <td style="width: 182px">
                                                                                            All Products<br />
                                                                                            <asp:ListBox ID="lbCSAvailableProducts" runat="server" DataSourceID="dsAllProducts"
                                                                                                DataTextField="productmodelnumber" DataValueField="productid" SelectionMode="Multiple"
                                                                                                Height="139px" Width="159px"></asp:ListBox>
                                                                                        </td>
                                                                                        <td style="width: 3px">
                                                                                            <asp:UpdatePanel ID="UpdatePanel2" runat="server" UpdateMode="conditional">
                                                                                                <ContentTemplate>
                                                                                                    <asp:Button ID="btnCrossSellMoveRight" runat="server" Text=">>>>" />
                                                                                                    <asp:Button ID="btnCrossSellMoveLeft" runat="server" Text="<<<<" /></ContentTemplate>
                                                                                            </asp:UpdatePanel>
                                                                                        </td>
                                                                                        <td>
                                                                                            Cross Sells<br />
                                                                                            <asp:ListBox ID="lbCrossSells" runat="server" DataSourceID="dsCrossSells" DataTextField="productmodelnumber"
                                                                                                DataValueField="Productid" SelectionMode="Multiple" Height="139px" Width="159px">
                                                                                            </asp:ListBox>
                                                                                            <asp:SqlDataSource ID="dsCrossSells" runat="server" CacheDuration="Infinite" EnableCaching="true"
                                                                                                ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>" SelectCommand="[MANAGERGetCrossSellsByProductID]"
                                                                                                SelectCommandType="StoredProcedure" UpdateCommand="[MANAGERUpdateCrossSells]"
                                                                                                UpdateCommandType="storedProcedure">
                                                                                                <UpdateParameters>
                                                                                                    <asp:Parameter Name="ProductID" Type="Int32" />
                                                                                                    <asp:Parameter Name="AssociatedProductID" Type="Int32" />
                                                                                                    <asp:Parameter Name="IsCrossSell" Type="boolean" />
                                                                                                </UpdateParameters>
                                                                                                <SelectParameters>
                                                                                                    <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="int32" />
                                                                                                </SelectParameters>
                                                                                            </asp:SqlDataSource>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                            </td>
                                                                        </tr>
                                                                    </table>
                                                                </ContentTemplate>
                                                            </asp:UpdatePanel>
                                                        </ContentTemplate>
                                                    </cc1:TabPanel>
                                                    <cc1:TabPanel ID="tpProductPromotion" runat="server" HeaderText="" Visible ="false" >
                                                        <HeaderTemplate>
                                                            
                                                        </HeaderTemplate>
                                                        <ContentTemplate>
                                                        
                                                        <asp:UpdatePanel ID="UpdatePanelProductPromotion" runat="server" UpdateMode="Conditional">
                                                                <ContentTemplate>
                                                                
                                                        <asp:Panel ID="pAddProductPromotion" runat="server" GroupingText ="Add Featured Location" >
                                                        <table border ="0" cellpadding ="0" cellspacing ="3"><tr><td>
                                                        Location Type:</td><td><asp:DropDownList ID="ddPromotionResourcetype"  runat="server" ><asp:ListItem Value ="11" Text ="Web Page"></asp:ListItem></asp:DropDownList></td><td>Location:<asp:DropDownList ID="ddPromotionResources"  runat="server" ><asp:ListItem Value ="1" Text ="CMS - Shopping Cart Page"></asp:ListItem></asp:DropDownList>
                                                        </td><td>Custom Description:<br />(Overrides product description)</td><td><asp:TextBox ID="txtPromotionDescription" Width ="400px" Height ="55px"  runat="server"  TextMode ="MultiLine" ></asp:TextBox>
                                                        </td><td><asp:Button ID="btnAddProductPromotion" runat="server" Text ="Add Product Promotion" OnClick="btnAddProductPromotion_Click" /></td></tr>
                                                        </table>
                                                        </asp:Panel>
                                                        <br /><br />
                                                        <asp:Panel ID="pCurrentPromotion" runat="server" GroupingText ="Current Featured Location(s)"  >
                                                            <asp:GridView ID="gvCurrentPromotions" Width ="100%" runat="server" AutoGenerateColumns="False" DataKeyNames="ProductPromotionID"
                                                                DataSourceID="dsProductPromotion">
                                                                <Columns>
                                                                    <asp:TemplateField ShowHeader="False">
                                                                        <EditItemTemplate>
                                                                            <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="True" CommandName="Update"
                                                                                Text="Update"></asp:LinkButton>
                                                                            <asp:LinkButton ID="LinkButton2" runat="server" CausesValidation="False" CommandName="Cancel"
                                                                                Text="Cancel"></asp:LinkButton>
                                                                        </EditItemTemplate>
                                                                        <ItemTemplate>
                                                                            <asp:LinkButton ID="LinkButton1" runat="server" CausesValidation="False" CommandName="Edit"
                                                                                Text="Edit"></asp:LinkButton>
                                                                            <asp:LinkButton ID="LinkButton2" runat="server" OnClientClick ="return confirm('Remove the product from the current location. Are you Sure?');" CausesValidation="False" CommandName="Delete"
                                                                                Text="Delete"></asp:LinkButton>
                                                                        </ItemTemplate>
                                                                    </asp:TemplateField>
                                                                    <asp:TemplateField HeaderText="Custom Description" HeaderStyle-Width ="400px" SortExpression="Description">
                                                                        <EditItemTemplate>
                                                                        <asp:TextBox ID="txtEditPromotionDescription" Text='<%# Bind("Description") %>' Width ="400px" Height ="75px"  runat="server"  TextMode ="MultiLine" ></asp:TextBox> 
                                                                        </EditItemTemplate>
                                                                        <ItemTemplate>
                                                                        <asp:Label ID="lblPromotionDescription" runat="server" Text='<%# Eval("Description") %>'></asp:Label>
                                                                        </ItemTemplate>
                                                                    </asp:TemplateField>
                                                                    <asp:TemplateField HeaderText="Feature Location" SortExpression="EmployeeName">
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="lblPromotedResourceName" runat="server" Text='<%# Eval("PromotedResourceName") %>'></asp:Label>
                                                                        </ItemTemplate>
                                                                    </asp:TemplateField>
                                                                    <asp:TemplateField HeaderText="Created By" SortExpression="CreatedByName">
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="Label4" runat="server" Text='<%# Eval("CreatedByName") %>'></asp:Label>
                                                                        </ItemTemplate>
                                                                    </asp:TemplateField>
                                                                    
                                                                    <asp:TemplateField HeaderText="Updated By" SortExpression="LastUpdatedByName">
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="lblLastUpdatedBy" runat="server" Text='<%# Eval("LastUpdatedByName") %>'></asp:Label>
                                                                        </ItemTemplate>
                                                                    </asp:TemplateField>
                                                                    <asp:TemplateField HeaderText="Last Updated" SortExpression="DateUpdated">
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="Label6" runat="server" Text='<%# Eval("DateUpdated") %>'></asp:Label>
                                                                        </ItemTemplate>
                                                                    </asp:TemplateField>
                                                                    <asp:TemplateField HeaderText="ProductPromotionID" InsertVisible="False" SortExpression="ProductPromotionID"
                                                                        Visible="False">
                                                                        <EditItemTemplate>
                                                                            <asp:Label ID="lblEditProductPromotionID" runat="server" Text='<%# Bind("ProductPromotionID") %>'></asp:Label>
                                                                        </EditItemTemplate>
                                                                        <ItemTemplate>
                                                                            <asp:Label ID="lblProductPromotionID" runat="server" Text='<%# Eval("ProductPromotionID") %>'></asp:Label>
                                                                        </ItemTemplate>
                                                                    </asp:TemplateField>
                                                                </Columns>
                                                            </asp:GridView>
                                                          
                                                        
                                                        </asp:Panel> 
                                                        </ContentTemplate>
                                                                </asp:UpdatePanel> 
                                                    </ContentTemplate> 
                                                    </cc1:TabPanel> 
                                                    
                                                </cc1:TabContainer>
                                            </td>
                                        </tr>
                                    </table>
                                    <!-- End Tabs -->
                                </td>
                            </tr>
                        </table>
            <%--        </td>
                </tr>
            </table>--%>
        </asp:Panel>
 <%--       <table width="100%" border="1">
            <tr>
                <td align="center">
                    <div id="updateProgressDiv" class="updateProgress" style="display: none">
                        <div style="margin-top: 0; text-align: center;">
                            <br />
                            <img src="../images/pleasewait.gif" alt="" /><br />
                            <span class="updateProgressMessage">Please Wait...</span>
                        </div>
                    </div>
                </td>
            </tr>
        </table>--%>
          <asp:SqlDataSource ID="dsProductPromotion" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                DeleteCommand="ManagerDeleteProductPromotion" DeleteCommandType="StoredProcedure"
                                                                InsertCommand="ManagerAddProductPromotion" InsertCommandType="StoredProcedure"
                                                                SelectCommand="ManagerGetProductPromotion" SelectCommandType="StoredProcedure"
                                                                UpdateCommand="ManagerUpdateProductPromotionText" UpdateCommandType="StoredProcedure">
                                                                <DeleteParameters>
                                                                    <asp:Parameter Name="ProductPromotionID" Type="Int32" />
                                                                </DeleteParameters>
                                                                <UpdateParameters>
                                                                    <asp:Parameter Name="ProductPromotionID" Type="Int32" />
                                                                    <asp:Parameter Name="Description" Type="String" />
                                                                    <asp:Parameter Name ="UpdatedBy" Type ="Int32" />
                                                                </UpdateParameters>
                                                                <SelectParameters>
                                                                    <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="Int32" />
                                                                </SelectParameters>
                                                                <InsertParameters>
                                                                    <asp:ControlParameter  ControlID ="TabContainer$tpProductPromotion$txtPromotionDescription" Name ="Description" DefaultValue =""  ConvertEmptyStringToNull="false" Type ="string" />
                                                                    <asp:ControlParameter  ControlID ="TabContainer$tpProductPromotion$ddPromotionResourcetype" Name ="ResourceTypeID" PropertyName ="SelectedValue" Type ="string" />
                                                                    <asp:ControlParameter  ControlID ="TabContainer$tpProductPromotion$ddPromotionResources" Name ="ResourceID" PropertyName ="SelectedValue" Type ="string" />
                                                                    <asp:QueryStringParameter Name ="ProductID" Type ="int32" QueryStringField ="ProductID" />
                                                                    <asp:Parameter Name="CreatedBy" Type="Int32" />
                                                                    <asp:Parameter name="ReturnValue"  type ="Int32"  Direction ="output" />
                                                                </InsertParameters>
                                                            </asp:SqlDataSource>
        <asp:SqlDataSource ID="dsAllProducts" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
            SelectCommand="GetAllProducts" SelectCommandType="StoredProcedure"></asp:SqlDataSource>
        <asp:SqlDataSource ID="dsAllRetailers" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
            SelectCommand="GetRetailers" SelectCommandType="StoredProcedure"></asp:SqlDataSource>
        <asp:SqlDataSource ID="dsProduct" runat="server" EnableCaching="false" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
            SelectCommand="[MANAGERGetProduct]" SelectCommandType="StoredProcedure" UpdateCommand="[MANAGERUpdateProduct]"
            UpdateCommandType="StoredProcedure" CancelSelectOnNullParameter="true">
            <SelectParameters>
                <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="Int32" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="ProductID" Type="Int32" />
                <asp:Parameter Name="ProductModelNumber" Type="String" />
                <asp:Parameter Name="Name" Type="String" />
                <asp:Parameter Name="Year" Type="Int32" />
                <asp:Parameter Name="Freight" Type="String" />
                <asp:Parameter Name="IsSupported" Type="Boolean" />
                <asp:Parameter Name="Active" Type="Boolean" />
                <asp:Parameter Name="ProductSeriesID" Type="Int32" />
                <asp:Parameter Name="IsActiveOnMktgSite" Type="Boolean" />
                <asp:Parameter Name="IsCurrentYearProduct" Type="Boolean" />
            </UpdateParameters>
        </asp:SqlDataSource>
        <asp:SqlDataSource ID="dsProductDescription" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
            SelectCommand="[MANAGERGetProduct]" SelectCommandType="StoredProcedure" UpdateCommand="[MANAGERUpdateProductDescription]"
            UpdateCommandType="StoredProcedure" CancelSelectOnNullParameter="true">
            <SelectParameters>
                <asp:QueryStringParameter Name="ProductID" QueryStringField="ProductID" Type="Int32" />
            </SelectParameters>
            <UpdateParameters>
                <asp:Parameter Name="ProductID" Type="Int32" />
                <asp:Parameter Name="Description" Type="String" />
            </UpdateParameters>
        </asp:SqlDataSource>
    
</asp:Content> 