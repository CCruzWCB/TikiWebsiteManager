<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ManageProducts.aspx.vb"
    Inherits="Manager_ManageProducts" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <title>Untitled Page</title>
</head>
<body style="margin-left: 0; margin-top: 0; margin-right: 0">
    <form id="form1" runat="server">
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
        <table border="0" cellpadding="3" cellspacing="0" width="100%">
            <tr>
                <td class="pageheader">
                    Manage Products</td>
            </tr>
        </table>
        <!-- Table Frame -->
        <table width="100%" border="0" cellpadding="5" cellspacing="0">
            <tr>
                <td style="width: 100%" align="center" valign="middle">
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
                            <td class="rcb_left">
                            </td>
                            <td class="rcb_mid">
                                <!-- Start Inner Panel Content -->
                                <!--Start Form Here-->
                                <table width="100%" border="0" cellpadding="4" cellspacing="0" >
                                    <tr>
                                        <td align="left" class="irowbold">
                                            Manage all Product Info:<br />
                                            <br />
                                            Add, Edit, Clone, Upload BOM (Parts), Upload Product Diagrams, Upload Product Images,
                                            Manage Product Specs, Upload Product Manuals, etc.
                                        </td>
                                    </tr>
                                </table>
                                <!-- END Inner Panel Content -->
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

                    <table width="100%" border="0" cellpadding="5" cellspacing="0">
                        <tr>
                            <td style="width: 100%" align="center" valign="middle">
                                <!--Start Form Here-->
                                <table width="100%" border="0" cellpadding="2" cellspacing="4">
                                    <tr>
                                        <td>
                                        </td>
                                    </tr>
                                    <tr>
                                        <td style="width: 100%" align="left">
                                            <asp:Button ID="btnAddNewProduct" runat="server" Text="Add New Product" CausesValidation="False"
                                                PostBackUrl="AddProduct.aspx" />
                                            &nbsp;&nbsp;<asp:Button ID="btnCloneProduct" runat="server" Text="Clone A Product"
                                                CausesValidation="False" PostBackUrl="CloneProduct.aspx" />
                                        </td>
                                    </tr>
                                    <tr>
                                        <td>
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <!--End Form Here-->
                    <!--End Form Here-->
                    <!-- Start TAB -->
                    <table style="height: 21px" border="0" cellpadding="0" cellspacing="0">
                        <tr>
                            <td>
                                <img src="../images/fake-tab_left.gif" width="3" height="21" alt="" /></td>
                            <td>
                                <img src="../images/fake-tab_repeat.gif" width="1" height="21" alt="" /></td>
                            <td style="height: 21px; white-space: nowrap; background-image: URL('../images/fake-tab_repeat.gif');"
                                class="tabtext">
                                &nbsp;Manage Products&nbsp;</td>
                            <td>
                                <img src="../images/fake-tab_right.gif" width="3" height="21" alt="" /></td>
                            <td style="background-image: URL('../images/fake-tab_topline.gif'); width: 100%;
                                height: 21px">
                            </td>
                        </tr>
                    </table>
                    <table cellpadding="0" cellspacing="0" border="0" width="100%">
                        <tr>
                            <td class="tabtablebody">
                                <table width="100%" border="0" cellspacing="0" cellpadding="5">
                                    <tr>
                                        <td><asp:Panel ID="pnlSelectModel" runat="server" DefaultButton="btnProductSearch">
                                            <!-- content goes here -->
                                            <table width="100%" border="0" cellpadding="4" cellspacing="0">
                                                <tr>
                                                    <td style="width: 100" align="CENTER" valign="middle" class="irowbold">
                                                        Enter Model Number to Manage OR Browse for Model below<br />
                                                        &nbsp;<asp:TextBox ID="txtModelNumber" runat="server"></asp:TextBox>
                                                        <asp:Button ID="btnProductSearch" runat="server" Text="Go" /></td>
                                                </tr>
                                                <tr>
                                                    <td style="width: 100" align="left" valign="middle" class="irowbold">
                                                        &nbsp;
                                                        
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="width: 100" align="left" valign="middle" class="irowbold">
                                                    </td>
                                                </tr>
                                            </table>
                                            </asp:Panel>
                                            <!-- end content here -->
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                    <!-- end TAB -->
                </td>
            </tr>
        </table>
    </form>
</body>
</html>
