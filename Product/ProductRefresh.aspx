<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="ProductRefresh.aspx.vb"
    Inherits="Manager_ProductRefresh" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server" >
<style>

#mytable {
	width: 700px;
	padding: 0;
	margin: 0;
}

caption {
	padding: 0 0 5px 0;
	width: 700px;	 
	font: italic 11px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;
	text-align: right;
}

th {
	font: bold 11px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;
	color: #4f6b72;
	border-right: 1px solid #C1DAD7;
	border-bottom: 1px solid #C1DAD7;
	border-top: 1px solid #C1DAD7;
	letter-spacing: 2px;
	text-transform: uppercase;
	text-align: left;
	padding: 6px 6px 6px 12px;
	background: #CAE8EA url(../images/bg_header.jpg) no-repeat;
}

th.nobg {
	border-top: 0;
	border-left: 0;
	border-right: 1px solid #C1DAD7;
	background: none;
}

td {
	border-right: 1px solid #C1DAD7;
	border-left: 1px solid #C1DAD7;
	border-bottom: 1px solid #C1DAD7;
	background: #fff;
	padding: 6px 6px 6px 12px;
	color: #4f6b72;
}


td.alt {
	background: #F5FAFA;
	color: #797268;
}

th.spec {
	border-left: 1px solid #C1DAD7;
	border-top: 0;
	background: #fff url(images/bullet1.gif) no-repeat;
	font: bold 10px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;
}

th.specalt {
	border-left: 1px solid #C1DAD7;
	border-top: 0;
	background: #f5fafa url(images/bullet2.gif) no-repeat;
	font: bold 10px "Trebuchet MS", Verdana, Arial, Helvetica, sans-serif;
	color: #797268;
}
 h4 
          {
              padding-bottom:5px;
            }
            
            .tag {font-weight:bold;padding-bottom:2px;}
            .value {padding-bottom:2px;padding-left:4px;color:#4097ca;font-size:14px;}
            
              
</style>



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
        <asp:ScriptManager ID="smManager" runat="server">
        </asp:ScriptManager>
         

         <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender3" runat="server"
            TargetControlID="uptimer">
        </cc1:UpdatePanelAnimationExtender>
        <cc1:UpdatePanelAnimationExtender ID="UpdatePanelAnimationExtender4" runat="server"
            TargetControlID="upwebstorerefreshbutton">
        </cc1:UpdatePanelAnimationExtender>
        
                 
        
          <asp:Panel ID="pnlForm" runat="server">
          
          <style>
         
              
          </style>
          <div style="width:980px;">

<div style="width:250px;border:1px solid #C1DAD7;margin-bottom:20px; text-align:left;background: #CAE8EA url(../images/bg_header.jpg) no-repeat;float:left;margin-right:20px;">
<div style="padding:5px;height:130px;">
<h4 style="text-align:center;">Refresh Parameters</h4>
                                           
<span class="tag">BizID:</span>  
<span class="value"><%= BizID %><br /></span>

<span class="tag">Public WebStoreID:</span>
<span class="value"><asp:Literal ID="Literal2" runat="server" Text="<%$ AppSettings:OrderMotionStoreIDPUBLIC%>" /><br /></span>

<span class="tag">Private WebStoreID:</span>  
<span class="value"><asp:Literal ID="Literal1" runat="server" Text="<%$ AppSettings:OrderMotionStoreIDPRIVATE%>" /><br /></span>
</div>
</div>                                            



                                                    
<asp:UpdatePanel ID="upWebStoreRefreshButton" runat="server" UpdateMode="Conditional">
    <ContentTemplate>

  
    <table id="tblRefreshOptions" style="width:700px;float:left;" >
    <thead>
        <tr>
            <th scope="col">Refresh Type</th>
            <th scope="col">Affected Sites</th>
            <th scope="col">Options</th>
            <th scope="col">Action</th>
        </tr>
        
    </thead>
    <tbody>

        <tr>
            <td>Webstore Refresh - PUBLIC</td>
            <td>Tiki Website (www.tikibrand.com)</td>
            <td><asp:CheckBox ID="chkUpdateCategoriesPublic" runat="server" Text="Update Categories" /></td>
            <td><asp:LinkButton ID="btnRefreshPublic" runat="server" Text="run now"  OnClick="btnRefresh_Click">
            <img src="../images/gears.png" height="20" />&nbsp;run now
            </asp:LinkButton></td>
        </tr>
          <tr>
            <td>Webstore Refresh - PRIVATE</td>
            <td>Tiki Website (sales.lamplight.com)</td>
            <td><asp:CheckBox ID="chkUpdateCategoriesPrivate" runat="server" Text="Update Categories" /></td>
            <td><asp:LinkButton ID="btnRefreshPrivate" runat="server" Text="run now"  OnClick="btnRefreshPrivate_Click">
            <img src="../images/gears.png" height="20" />&nbsp;run now
            </asp:LinkButton></td>
        </tr>
        <tr>
            <td>Webstore Refresh - LL Rep Portal</td>
            <td>Rep Portal website (repsportal.tikibrand.com)</td>
            <td>&nbsp;</td>
            <td><asp:LinkButton ID="btnRefreshRepPortal" runat="server" Text="run now"  OnClick="btnRefreshRepPortal_Click">
            <img src="../images/gears.png" height="20" />&nbsp;run now
            </asp:LinkButton></td>
        </tr>
           <tr>
            <td>Shipping Rate Refresh</td>
            <td>All Sites</td>
            <td>&nbsp;</td>
            <td><asp:LinkButton ID="btnRefreshShippingRates" runat="server" Text="run now"  OnClick="btnRefreshShippingRates_Click">
            <img src="../images/gears.png" height="20" />&nbsp;run now
            </asp:LinkButton></td>
        </tr>
        </tbody>


    </table>
                                                                
     <br />
                                                            </ContentTemplate>
</asp:UpdatePanel>
</div>

                                                        <div style="float:left;width:100%;">

                                                        <asp:Timer ID="tRefreshTimer" runat="server" Interval="10000" OnTick="CheckStatus">
                                                        </asp:Timer>
                                                        <asp:UpdatePanel ID="upTimer" runat="server" UpdateMode="Conditional" >
                                                            <Triggers>
                                                                <asp:AsyncPostBackTrigger ControlID="tRefreshTimer" />
                                                            </Triggers>
                                                            <ContentTemplate><br /><br /><asp:HiddenField ID="hdnRefreshLogID" runat="server" />
                                                                <img id="imgProcessing" src="../Images/processing.gif" runat="server" visible="false" /><br />
                                                                <span style="font: italic 11px 'Trebuchet MS, Verdana, Arial, Helvetica, sans-serif';font-size:14px">
                                                                Current Status:  <asp:Label ID="lblStatus" runat="server"  ForeColor="#4097ca" Font-Bold="True"></asp:Label>
                                                                </span>

                                                                
                                                                <br /><br />
                                                                <asp:GridView ID="gvRefreshLog" style="min-width:900px;max-width:85%;" runat="server" AutoGenerateColumns="False"  DataSourceID="dsRefreshLog">
                                                                    <Columns>
                                                                        
                                                                        <asp:TemplateField ItemStyle-HorizontalAlign="left" ItemStyle-Width="415" HeaderText="Refresh Type" SortExpression="RefreshType" >
                                                                             
                                                                             <ItemTemplate>
                                                                                <%# Eval("RefreshType")%><br />
                                                                                <%# Eval("StatusNotes")%>
                                                                             </ItemTemplate>
                                                                             
                                                                        </asp:TemplateField>
                                                                        <%--<asp:BoundField DataField="RefreshItem" HeaderText="Item" SortExpression="RefreshType" />--%>
                                                                        <asp:BoundField DataField="StartDate" ItemStyle-Width="140" HeaderText="Start Date" SortExpression="StartDate" />
                                                                        <asp:BoundField DataField="EndDate" ItemStyle-Width="140" HeaderText="End Date" SortExpression="EndDate" />
                                                                        <%--<asp:BoundField DataField="CreatedByName" ItemStyle-Width="90" HeaderText="Created By" ReadOnly="True" SortExpression="CreatedByName" />--%>
                                                                        <asp:BoundField DataField="Status" HeaderText="Status" ItemStyle-Width="65" SortExpression="Status" />
                                                                        
                                                                    </Columns>
                                                                </asp:GridView>
                                                                <asp:SqlDataSource ID="dsRefreshLog" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                    SelectCommand="GetRefreshLog_Top50" SelectCommandType="StoredProcedure">
                                                                    <SelectParameters>
                                                                    <%--<asp:Parameter Name="RefreshType" Type="string" DefaultValue="Web Store" />--%>
                                                                    </SelectParameters>
                                                                    
                                                                    </asp:SqlDataSource>
                                                                    <asp:SqlDataSource ID="dsRefreshStatus" runat="server" ConnectionString="<%$ ConnectionStrings:BaseDBConnection %>"
                                                                    SelectCommand="GetRefreshStatus" SelectCommandType="StoredProcedure" ></asp:SqlDataSource>
                                                            </ContentTemplate>
                                                        </asp:UpdatePanel>
                                      
        <!-- Table Frame -->
      </div>
        </asp:Panel>
        <table width="100%" border="1">
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
        </table>

    
</asp:Content>