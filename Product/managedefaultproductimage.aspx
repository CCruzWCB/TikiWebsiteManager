<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false"
    CodeFile="managedefaultproductimage.aspx.vb" Inherits="Product_managedefaultproductimage"
    Title="Untitled Page" %>
<%@ Import Namespace ="Management" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" runat="Server">
    &nbsp;&nbsp;<b> Web Site Default Product Image (Image not available) </b>
    &nbsp;&nbsp;&nbsp;
    <br />
    <br />
    <asp:Panel ID="pPrimary" runat="server" Visible="false">
        <div style="border: 1px solid black;">
        <table border="0">
            <tr>
                <td>&nbsp;
                    <asp:Image ID="imgPrimary" runat="server" /></td>
                <td valign ="top"  align ="center" >
                    <asp:CheckBox ID="chkAutoResize" Checked ="true"  runat="server" Text="Enable auto resize" ToolTip="This feature allows for automatic creation of thumbnail and regular image size based on image properites. " AutoPostBack="True" />
                    <asp:Panel ID="pImageSizeSettings" runat="server" Visible ="true" >
                        <table border="0" bgcolor="#eeeeee" width ="350">
                            <tr><td colspan ="3" align="center" ><b>Image Resize Properites</b></td></tr>
                            <tr><td>&nbsp;</td><td>Width</td><td>Height</td></tr>
                            <tr>
                                <td>
                                    Feature:</td>
                                <td >
                                    
                                    <asp:label ID="lblFeatureImageSizeWidth" runat="server" ></asp:label>
                                    
                                   </td><td>
                                    <asp:label ID="lblFeatureImageSizeHeight" runat="server" ></asp:label>
                                    </td> 
                                    
                            </tr>
                            <tr>
                                <td>
                                    Small:</td>
                                <td >
                                    
                                    <asp:label ID="lblSmallImageSizeWidth" runat="server" ></asp:label>
                                    
</td><td>
                                    <asp:label ID="lblSmallImageSizeHeight" runat="server" ></asp:label>
         </td>                            
                            </tr>
                            <tr>
                                <td>
                                    Regular:
                                </td>
                                <td >
                                    <asp:label ID="lblRegularImageSizeWidth" runat="server" ></asp:label>
                                    </td><td>
                                    <asp:label  ID="lblRegularImageSizeHeight" runat="server" ></asp:label></td>
                                    
                            </tr>
                            <tr>
                                <td>
                                    Large:
                                </td>
                                <td >
                                    
                                    <asp:label ID="lblLargeImageSizeWidth" runat="server" ></asp:label>
                                    </td><td>
                                    <asp:label ID="lblLargeImageSizeHeight" runat="server" ></asp:label></td>
                            </tr>
                            <tr>
                        <td colspan="3" align="center"><br />
                        <asp:FileUpload ID="FileProductImageUpload" runat="server" /><br /><br />
                            <asp:Button ID="btnUpdateImage" runat="server" Text="Update Image" OnClientClick="return  confirm('This will replace all existing product image sizes. Are you sure');" />
                            <asp:Button ID="btnRegenerateImage" runat="server" Text="Reload Default Images" OnClientClick="return  confirm('This will replace all existing product image sizes. Are you sure');" ToolTip="This option generate resized image based on the current largest image specified." />
                            <br /><br />
                            </td>
            </tr></table> 
                    </asp:Panel>
                </td>
            </tr>
            
            <tr>
                <td colspan="2" align="center">
                    <!-- 3 Image Table !-->
                    <asp:Panel ID="pProductImages" runat="server" Visible ="false"  >
                    <table>
                        <tr>
                            <td>
                                    <table>
                                    <tr>
                                        <td align=center >Featured Image</td>
                                    </tr>
                                    <tr>
                                        <td align=center >
                                <asp:Image ID="imgFeature" runat="server" />        
                                         </td>
                                   </tr>
                                   <tr>
                                   <td>
                                   <asp:FileUpload ID="FileFeatureImageUpload" runat="server" Visible="False" />
                                <br /><asp:Button ID="btnUpdateFeatureImage" runat="server" Text="Update Image" Visible="False" />
                                   </td>
                                   </tr>
                                </table>
                                <br />
                                <table>
                                    <tr>
                                        <td align=center >Small Image</td>
                                    </tr>
                                    <tr>
                                        <td align=center >
                                <asp:Image ID="imgSmall" runat="server" />        
                                         </td>
                                   </tr>
                                   <tr>
                                   <td style="height: 67px">
                                   <asp:FileUpload ID="FileSmallImageUpload" runat="server" Visible="False" />
                                <br /><asp:Button ID="btnUpdateSmallImage" runat="server" Text="Update Image" Visible="False" />
                                   </td>
                                   </tr>
                                </table>
                            </td>
                            <td>
                                    <table>
                                    <tr>
                                        <td align=center >Regular Image </td>
                                    </tr>
                                    <tr>
                                        <td align=center ><asp:Image ID="imgRegular" runat="server" />
                                        </td>
                                   </tr>
                                   <tr>
                                   <td >
                                <asp:FileUpload ID="FileRegularImageUpload" runat="server" Visible="False" />
                                </td>
                                   </tr>
                                   <tr>
                                   <td><asp:Button ID="btnUpdateRegularImage" runat="server" Text="Update Image" Visible="False" />
                                </td>
                                   </tr>
                                </table>
                                </td>
                            <td>
                             <table>
                                    <tr>
                                        <td align=center >Large Image</td>
                                    </tr>
                                <tr><td align=center >
                                <asp:Image ID="imgLarge" runat="server" />
                                </td>
                                <tr><td>
                                <asp:FileUpload ID="FileLargeImageUpload" runat="server" Visible="False" />
                                <br /><asp:Button ID="btnUpdateLargeImage" runat="server" Text="Update Image" Visible="False" />
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
        <asp:FileUpload ID="FileUploadAlternate" runat="server" />
        <asp:Button ID="btnAddAlternateImage" runat="server" Text="Add Image" />
        <asp:GridView ID="gvAlternateImages" runat="server" AutoGenerateColumns="False" DataKeyNames="ImageID"
            DataSourceID="dsAlternateImages" AllowPaging="True" AutoGenerateEditButton="True">
            <Columns>
                <asp:TemplateField ShowHeader="False">
                <EditItemTemplate>
                        <asp:Label ID ="lblEditImageID" runat="server"  Text ='<%# Eval("ImageID") %>'></asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:LinkButton ID="LinkButton1" OnClientClick ="return confirm('You are about to delete this image. Are you sure?');" runat="server" CausesValidation="False" CommandName="Delete"
                         CommandArgument ='<%# Eval("ImageID") %>'    Text="Delete"></asp:LinkButton>
                    </ItemTemplate>
                </asp:TemplateField>
                <asp:BoundField DataField="ImageID" Visible ="False"  HeaderText="ImageID" ReadOnly="True"
                    SortExpression="ImageID" />
                <asp:TemplateField HeaderText="Name" SortExpression="Name">
                    <EditItemTemplate>
                        <asp:Label id="lblEditName" runat="server" Text='<%# Bind("Name") %>'> </asp:Label>
                    </EditItemTemplate>
                    <ItemTemplate>
                        <asp:Image AlternateText ='<%# Eval("Name") %>' ImageAlign ="Middle" runat="server" ImageUrl = '<%# Utilities.GetAlternateThumbnailImagePath (Eval("ImagePath")) %>' />
                        
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
                <asp:TemplateField HeaderText="ImagePath" SortExpression="ImagePath" Visible ="False" >
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
            SelectCommand="SELECT Image.ImageID, Image.Name, Image.Description, Image.ImagePath, Image.DateCreated FROM Image INNER JOIN ImageResourceJUNC ON Image.ImageID = ImageResourceJUNC.ImageID WHERE (ImageResourceJUNC.ResourceTypeID = 3) AND (ImageResourceJUNC.ResourceID = @ProductSeriesID)" DeleteCommand="DeleteImage" DeleteCommandType="StoredProcedure" UpdateCommand="Update [Image] Set Description = @Description Where ImageID = @ImageID">
            <SelectParameters>
                <asp:QueryStringParameter Name="ProductSeriesID" QueryStringField="ProductSeriesID" />
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
        <div id="divImagePreview" style="display:none;">
    <img id="imgPreview" alt ="" src ="" />
    </div>

</asp:Content>
