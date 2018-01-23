<%@ Register Src="../ClientWebControls/ClientDataCube.ascx" TagName="ClientDataCube"   TagPrefix="uc2" %>
<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="manageProductlistings.aspx.vb" Inherits="ContentManagement_manageProductlistings" title="Untitled Page" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
<table cellspacing="6" cellpadding="0" border="0">
								<tbody>
									<tr>
										<TD width="200" height="200">
                                            <uc2:ClientDataCube ID="ClientDataCube1" EditMode="True" DataCubeID="DataCubePL1" TextBody="[Click to Edit DataCube 1] "
                                             Width = "126" Height ="126"
												Target="_SELF" BorderSize="5"
												BorderColor="yellow" TextColor="#000000" EnableBorder="True" runat="server" />
                                        </TD>
									
									</tr>
																	</tbody>
							</table>
</asp:Content>

