<%@ Page Language="vb" AutoEventWireup="false" CodeFile="mSelectMergeFields.aspx.vb" Inherits="DataCube_mSelectMergeFields" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD runat="server" >
		<title>Select Data Fields</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="JavaScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<script type="text/javascript"  language="javascript">
<!--
function UpdateMasterForm()
{
	var opener = window.opener;
	if (Form1.hdnColumn.value != "")
	{
		try
		{
			var control = eval("opener.document.forms[0]." + Form1.hdnControlName.value)
			control.value += " " + Form1.hdnColumn.value 
		}
		catch(e)
		{
			alert("This option is not enabled for this control type [" + control.tagName + "!")
		}
	}
}


function ExitWindow()
{
	var oElement = null
	window.returnValue = oElement
	window.close();
	
}
//-->
		</script>
	</HEAD>
	<body onload="UpdateMasterForm();">
		<form id="Form1" method="post" target="_self" runat="server">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td colspan="3">To Merge Data Fields, First, highlight textarea on main form and 
						then double click a field below.
					</td>
				</tr>
				<tr>
					<td colspan="3"><br>
					</td>
				</tr>
				<tr>
					<td colspan="3"><br>
						<asp:DataGrid id="dgFieldName" runat="server" BorderColor="#DEBA84" BorderStyle="None" CellSpacing="2"
							BorderWidth="1px" BackColor="#DEBA84" CellPadding="3">
							<SelectedItemStyle Font-Bold="True" ForeColor="White" BackColor="#738A9C"></SelectedItemStyle>
							<ItemStyle ForeColor="#8C4510" BackColor="#FFF7E7"></ItemStyle>
							<HeaderStyle Font-Bold="True" ForeColor="White" BackColor="#A55129"></HeaderStyle>
							<FooterStyle ForeColor="#8C4510" BackColor="#F7DFB5"></FooterStyle>
							<Columns>
								<asp:ButtonColumn Text="Add" CommandName="add"></asp:ButtonColumn>
							</Columns>
							<PagerStyle HorizontalAlign="Center" ForeColor="#8C4510" Mode="NumericPages"></PagerStyle>
						</asp:DataGrid>
					</td>
				</tr>
				<tr>
					<td align="center" colspan="3"><INPUT id="btnCancel" onclick="ExitWindow();" type="button" value="Close" /></td>
				</tr>
			</table>
			<input type="hidden" id="hdnColumn" runat="server"> <input type="hidden" id="hdnControlName" runat="server" NAME="hdnControlName">
		</form>
	</body>
</HTML>
