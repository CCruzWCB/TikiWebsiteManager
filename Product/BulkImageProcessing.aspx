<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="BulkImageProcessing.aspx.vb" Inherits="Product_BulkImageProcessing" title="Untitled Page" %>
<%@ MasterType VirtualPath="~/MasterPages/MasterPage.master" %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">
    To import images for multiple products, please follow the steps below.

<table><tr><td align="left" >
<ul>
<li>
1) Copy product images to the bulk image staging directory shown below. 
<br />To open directory folder <a href='<%=ConfigurationManager.AppSettings("ImageBulkUploadDirectory").ToString%>'>click here</a> or navigate to the following network path below.<br /><br />
<strong>Bulk Image Staging Directory: </strong><%=ConfigurationManager.AppSettings("ImageBulkUploadDirectory").ToString%>
<br /><br /><strong>Image Sizing</strong>: All images will be optimized and resized according to the settings shown below. To prevent image distortion (squished or stretched like appearance), all images should be "square" or equally proportion in width & height and have a minumum height and/or width of 400 pixels. 
           <table border="0" width ="350" class ="box">
                            <tr><td class="boxheader" colspan ="3" align="center" ><b>Image Resize Properites</b></td></tr>
                            <tr><td>&nbsp;</td><td>Width</td><td>Height</td></tr>
                            <tr>
                                <td>
                                    Featured/Alt. Thumbnail:</td>
                                <td>
                                    <asp:Literal runat ="server" Text ='<%$ AppSettings:MaxFeatureWidth %>'></asp:Literal>
                                    
                                    
                                   </td><td>
                                   <asp:Literal ID="Literal1" runat ="server" Text ='<%$ AppSettings:MaxFeatureHeight %>'></asp:Literal>
                                    
                                    </td> 
                                    
                            </tr>
                            <tr>
                                <td>
                                    Small:</td>
                                <td >
                                <asp:Literal ID="Literal2" runat ="server" Text ='<%$ AppSettings:MaxSmallWidth %>'></asp:Literal>

                                    
</td><td>
<asp:Literal ID="Literal3" runat ="server" Text ='<%$ AppSettings:MaxSmallHeight %>'></asp:Literal>
                                    
         </td>                            
                            </tr>
                            <tr>
                                <td>
                                    Regular:
                                </td>
                                <td >
                                    <asp:Literal ID="Literal4" runat ="server" Text ='<%$ AppSettings:MaxRegularWidth %>'></asp:Literal>
                                    </td><td>
                                    <asp:Literal ID="Literal5" runat ="server" Text ='<%$ AppSettings:MaxRegularHeight %>'></asp:Literal></td>
                                    
                            </tr>
                            <tr>
                                <td>
                                    Large:
                                </td>
                                <td >
                                    
                                    <asp:Literal ID="Literal6" runat ="server" Text ='<%$ AppSettings:MaxLargeWidth %>'></asp:Literal>
                                    </td><td>
                                    <asp:Literal ID="Literal7" runat ="server" Text ='<%$ AppSettings:MaxLargeHeight %>'></asp:Literal></td>
                            </tr>
                           </table> 
<br /><br /><strong>File Naming convension</strong>: Primary product image should match product model/item number of the product. For example,  Item # G5478 should be name "G5478.jpg". Existing product images will be replaced by the newly uploaded file. 
<br /><br /><strong>Additional Images</strong>: If uploading additional product image views, file names should contain the based product model/item number followed by an understore and description (i.e. "G5478_sideview.jpg"). Note: additional images file are renamed and appended to the product. To add/remove additional images from the product, select the <a  href ="selectproduct.aspx">Manage Product Images</a>  option for the Product Menu.

Note: Access to the network path is based window's user credentials. Please contact the <a href="mailto:helpdesk@wcbradley.com?subject=Request Access to Network Directory&body=Please grant me access to copy files to the following directory: <%=ConfigurationManager.AppSettings("ImageBulkUploadDirectory").ToString%>">help desk</a> if you experience problems accessing the network path. 
<br /><br />
</li>
<li>2) The image processing service is scheduled to check for files every 30 minutes. So, Please allow up to 30 minutes for the BIP service to detect and process new files. Note: files will be removed from the directory after they have been processed successfully.
<br /><br />
</li>


</ul>


</td></tr></table>

                                
</asp:Content>

