<%@ Page Language="VB" AutoEventWireup="false" MasterPageFile="~/MasterPages/MasterPage.master"  CodeFile="ManageCurrentYearProduct.aspx.vb" Inherits="Product_ManageCurrentYearProduct" %>

<%@ Register Assembly="System.Web.Extensions, Version=1.0.61025.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35"
    Namespace="System.Web.UI" TagPrefix="asp" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
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

    <asp:ScriptManager ID="ScriptManager1" runat="server">
    </asp:ScriptManager>
    
    <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender1" runat="server"
        TargetControlID="UpdatePanel1">
        <Animations>
                        <OnUpdating>
                        <Parallel duration="0">
                            <%-- place the update progress div over the gridview control --%>
                            <ScriptAction Script="onUpdating();" />  
                         </Parallel>
                    </OnUpdating>
                    <OnUpdated>
                        <Parallel duration="0">
                            <%--find the update progress div and place it over the gridview control--%>
                            <ScriptAction Script="onUpdated();" /> 
                        </Parallel> 
                    </OnUpdated>

                                      
        </Animations>
    </cc1:UpdatePanelAnimationExtender>
    <asp:Panel ID="pnlForm" runat="server">
        <asp:UpdatePanel ID="UpdatePanel1" runat="server">
            <ContentTemplate>
                <table>
                   
                    <tr>
                        <td style="width: 182px">All Products<br />
                            <asp:ListBox ID="lbNonMktg" runat="server" DataSourceID="dsNonMktgProducts" DataTextField="productmodelnumber"
                                DataValueField="productid" SelectionMode="Multiple" Height="139px" Width="159px">
                            </asp:ListBox>
                            <asp:SqlDataSource ID="dsMktgProducts" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                SelectCommand="MANAGERGetAllCurrentYearProducts" SelectCommandType="StoredProcedure">
                            </asp:SqlDataSource>
                        </td>
                        <td style="width: 3px">
                            <asp:Button ID="btnMoveRight" runat="server" Text=">>>>" />
                            <asp:Button ID="btnMoveLeft" runat="server" Text="<<<<" /></td>
                        <td>Current Year Products<br />
                            <asp:ListBox ID="lbMktg" runat="server" DataSourceID="dsMktgProducts" DataTextField="productmodelnumber"
                                DataValueField="productid" SelectionMode="Multiple" Height="139px" Width="159px">
                            </asp:ListBox>
                            <asp:SqlDataSource ID="dsNonMktgProducts" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                SelectCommand="MANAGERGetAllNONCurrentYearProducts" SelectCommandType="StoredProcedure">
                            </asp:SqlDataSource>
                        </td>
                    </tr>
                 </table>
            </ContentTemplate>
        </asp:UpdatePanel>
    </asp:Panel>
    <table width="100%" border="1">
        <tr>
            <td align="center">
                <div id="updateProgressDiv" class="updateProgress" style="display: none">
                    <div style="margin-top: 0; text-align: center;"><br />
                        <img src="../images/pleasewait.gif" alt="" /><br />
                        <span class="updateProgressMessage">Please Wait...</span>
                    </div>
                </div>
            </td>
        </tr>
    </table>
</asp:Content>
