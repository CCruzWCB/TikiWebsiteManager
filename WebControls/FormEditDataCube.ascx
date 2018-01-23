<%@ Control Language="vb" AutoEventWireup="false" CodeFile="FormEditDataCube.ascx.vb" Inherits="WebControls_FormEditDataCube" TargetSchema="http://schemas.microsoft.com/intellisense/ie5"  %>
<%@ Register Src="../ClientWebControls/ClientDataCube.ascx" TagName="ClientDataCube"
    TagPrefix="uc2" %>
    <!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN" >


<table>
	<tr>
		<TD align="center">
			<asp:Panel id="PanelEditMode" runat="server">
				<asp:Label id="lblMessage" Runat="server"></asp:Label>
				<asp:table id="TableDataCube" runat="server" Width="401px" BorderWidth="3px" BorderStyle="Solid">
					<asp:TableRow ID="trHeader">
						<asp:TableCell ID="tdEdit">Edit:</asp:TableCell>
						<asp:TableCell ID="tdHeader"></asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trType">
						<asp:TableCell>Type:</asp:TableCell>
						<asp:TableCell ID="tdType">
							<asp:DropDownList id="lstType" AutoPostBack="true" runat="server"></asp:DropDownList>
							<asp:RequiredFieldValidator id="rfvType" runat="server" ErrorMessage="Select a DataCube Type." ControlToValidate="lstType"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trElement">
						<asp:TableCell ID="tdElementName">Element Name Here: </asp:TableCell>
						<asp:TableCell ID="tdElement">
							<asp:TextBox id="txtElement" Columns="25" runat="server"></asp:TextBox>
							<asp:Button id="btnBrowseElements" CausesValidation="False" runat="server" Text="..."></asp:Button>
							<asp:RequiredFieldValidator id="rfvElement" runat="server" ErrorMessage="Please Select an Element." ControlToValidate="txtElement"
								InitialValue="Enter or Select Element"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trDescription">
						<asp:TableCell>Description: </asp:TableCell>
						<asp:TableCell ID="tdDescription">
							<asp:Image id="imgMergeDescription" runat="server" ImageUrl ="~/images/icons/wlink.gif"></asp:Image>
							<asp:TextBox id="txtDescription" runat="server" Rows="5" Columns="35" TextMode="MultiLine"></asp:TextBox>
							<asp:RequiredFieldValidator id="rfvDescription" runat="server" ErrorMessage="Description is Required!" ControlToValidate="txtDescription"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trTargetUrl">
						<asp:TableCell>Base URL: </asp:TableCell>
						<asp:TableCell ID="tdTargetUrl">
							<asp:TextBox id="txtTargetUrl" MaxLength="255" runat="server"></asp:TextBox>
							<asp:RequiredFieldValidator id="rfvTargetUrl" runat="server" ErrorMessage="Please enter the target Url." ControlToValidate="txtTargetUrl"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trTargetPage">
						<asp:TableCell ID="tdTargetPageHeader">Target Page:</asp:TableCell>
						<asp:TableCell ID="tdtargetPage">
							<asp:TextBox id="txtTargetPage" runat="server" MaxLength="255"></asp:TextBox>
							<asp:RequiredFieldValidator id="rfvTargetPage" runat="server" ErrorMessage="Please enter the name of the target Page( i.e. default.aspx)."
								ControlToValidate="txtTargetPage"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trTarget">
						<asp:TableCell>Page Target:</asp:TableCell>
						<asp:TableCell ID="tdTarget">
							<asp:DropDownList id="lstTarget" runat="server">
								<asp:ListItem Value="0">Select a Target</asp:ListItem>
								<asp:ListItem Value="_BLANK">New Window</asp:ListItem>
								<asp:ListItem Value="_SELF">Current Window</asp:ListItem>
							</asp:DropDownList>
							<asp:RequiredFieldValidator id="rfvTarget" runat="server" ErrorMessage="Select a target for Url." ControlToValidate="lstTarget"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trOptionalText">
						<asp:TableCell>Optional Text:</asp:TableCell>
						<asp:TableCell ID="tdOptionalText">
							<asp:Image id="ImageMergeOptionalText" runat="server" ImageUrl="images/wlink.gif"></asp:Image>
							<asp:TextBox id="txtOptionalText" runat="server"></asp:TextBox>
							<asp:RequiredFieldValidator id="rfvOptionalText" runat="server" ErrorMessage="Description is Required!" ControlToValidate="txtOptionalText"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trSourceCode">
						<asp:TableCell>Source Code: </asp:TableCell>
						<asp:TableCell ID="tdSourceCode">
							<asp:TextBox id="txtSourceCode" runat="server"></asp:TextBox>
							<asp:RequiredFieldValidator id="rfvSourceCode" runat="server" ErrorMessage="Source Code is Required!" ControlToValidate="txtSourceCode"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trDisplayImage">
						<asp:TableCell ID="tdImage">Current Image: </asp:TableCell>
						<asp:TableCell>
							<asp:Image id="ElementImage" runat="server"></asp:Image>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trDisplayFlash">
						<asp:TableCell ID="tdCurrentFlash">Current Flash: </asp:TableCell>
						<asp:TableCell>
							<div id="divFlashContainer" runat="server" ></div>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trImage">
						<asp:TableCell>New Image:</asp:TableCell>
						<asp:TableCell>
							<INPUT id="ImageFile" type="file" runat="server" NAME="ImageFile">
							<asp:RequiredFieldValidator id="rfvImageFile" runat="server" ErrorMessage="New Image Required!" ControlToValidate="ImageFile"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trFlashFile">
						<asp:TableCell>New Flash File:</asp:TableCell>
						<asp:TableCell ID="tdFlashFile">
							<INPUT id="FlashFile" type="file" runat="server" NAME="FlashFile">
							<asp:RequiredFieldValidator id="rfvFlashFile" runat="server" ErrorMessage="New Flash file Required!" ControlToValidate="FlashFile"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trImagePath">
						<asp:TableCell>Image Path:</asp:TableCell>
						<asp:TableCell ID="tdImagePath">
							<asp:TextBox id="txtImagePath" runat="server"></asp:TextBox>
							<asp:RequiredFieldValidator id="rfvImagePath" runat="server" ErrorMessage="New Image Path!" ControlToValidate="txtImagePath"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trWidth">
						<asp:TableCell>Width: </asp:TableCell>
						<asp:TableCell ID="tdWidth">
							<asp:TextBox id="txtWidth" runat="server" Columns="3"></asp:TextBox>
							<asp:RequiredFieldValidator id="rfvWidth" runat="server" ErrorMessage="New Image Width Required." ControlToValidate="txtWidth"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trHeight">
						<asp:TableCell>Height: </asp:TableCell>
						<asp:TableCell ID="tdHeight">
							<asp:TextBox id="txtHeight" runat="server" Columns="3"></asp:TextBox>
							<asp:RequiredFieldValidator id="rfvHeight" runat="server" ErrorMessage="New Image Height Required." ControlToValidate="txtHeight"></asp:RequiredFieldValidator>
						</asp:TableCell>
					</asp:TableRow>
					<asp:TableRow ID="trEnable">
						<asp:TableCell>Enable</asp:TableCell>
						<asp:TableCell ID="tdEnable">
							<asp:DropDownList id="lstEnable" runat="server">
								<asp:ListItem Value="1">Enabled</asp:ListItem>
								<asp:ListItem Value="0">Disabled</asp:ListItem>
							</asp:DropDownList>
						</asp:TableCell>
					</asp:TableRow>
				</asp:table>
			</asp:Panel>
		</TD>
	</tr>
	<tr>
		<td><asp:Panel id="PanelConfirm" Visible="false" runat="server">
				<table width="400">
					<tr>
						<TD align="center">Please Test and Confirm DataCube Below. Click Image or Text to 
							preview any hyperlinks.
						</TD>
					</tr>
					<tr>
						<TD align="center">
							<uc2:ClientDataCube ID="DataCubePreview" Enabled="False" EnableBorder="True" BorderSize="1"
								BorderColor="black" runat="server" />
                            <uc2:ClientDataCube ID="ClientDataCube1" runat="server" />
								
                        </TD>
					</tr>
				</table>
			</asp:Panel></td>
	</tr>
	<tr>
		<TD align="center" style="height: 26px">
			<asp:button id="btnCancel" runat="server" Text="Cancel" CausesValidation="False"></asp:button>
			<asp:button id="btnConfirm" runat="server" Text="Confirm"></asp:button>
			<asp:button id="btnSave" runat="server" Text="Save"></asp:button>
		</TD>
	</tr>
</table>
<INPUT id="hdnDataCubeID" type="hidden" name="hdnDataCubeID" runat="server">
