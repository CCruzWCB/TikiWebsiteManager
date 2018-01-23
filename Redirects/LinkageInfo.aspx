<%@ Page Language="VB" MasterPageFile="~/MasterPages/MasterPage.master" AutoEventWireup="false" CodeFile="LinkageInfo.aspx.vb" Inherits="Redirects_LinkageInfo" title="Untitled Page" %>
<%@ MasterType VirtualPath ="~/MasterPages/MasterPage.master"  %>
<asp:Content ID="Content1" ContentPlaceHolderID="ContentMain" Runat="Server">

<div style="width:70%;margin:5px;" class ="box">
<div style="margin:5px;" class ="graybox" ><b>How to use "GoToPage.aspx" redirector page.</b></div>

<div style ="text-align :left ;">
Use the page and querystring values below to provide linkage to pages within the website from email campaigns or other external sources. Using the links below will ensure 
user's origination are tracked and then navigated to the requested resource within the website. 

<br />
<br />

<strong>Tracking Guide: </strong><br />

The gotoPage allows for two additional tracking parameters. The first tracking parameter is called "SourceCode". This optional querystring parameter
is used to track the external source (email, externally linked site, etc...) from which the user originated. The second optional tracking
parameter is called "PLID". This parameter is used in conjunction with the Sourcecode to set the pricelist using the external pricelist code (which is generated on pricelist creation). 
<br /><br />
To apply tracking, append querystring values to existing GoToPage parameters. 
<BR />For example, to track a promotion from sizzleontheGrill.com,  the url would appear like: <br />
<br />
http://www.charbroil.com/gotopage.aspx?URLType=7&Value=21&<font color=red>SourceCode=SizzleontheGrill&PLID=SIZZ58H8D49</font>


<br />

<br />

<strong>Event Tracking via <a target ="_blank" title ="Go to Google Analytics" href ="https://www.google.com/analytics/">Google Analytics</a> : </strong><br />
The gotoPage allows for the capturing and reporting of Google Analytic tracking events. To capture Events the following parameters must be used: <br />
<br />
EC - Defines the Event Category. This String URL parameter value is <b>required</b>. <br />
EA - Defines the Event Action. For example, "Click", "Play", etc... This String URL parameter value is <b>required</b>.<br /> 
EL - Defines the Event Label. This String URL parameter value is optional. <br />
EV - Defines the Event Value. This Numeric URL parameter value is optional. Note: this value is not tracked uniquely.  <br />
<br />
<u>Example</u>
<BR />To track a specific product clicked from an Ad located on sizzleontheGrill.com,  the url would be: <br />
<br />
http://www.charbroil.com/gotopage.aspx?URLType=7&Value=21&<font color=red>EC=Sizzle+On+the+Grill&EA=CB+Aug+Ad2010&EL=463666510&EV=1</font>
<Br /><Br />
<font color="red">Note: all url parameters should be safe web characters and not contain special characters unless they have been encoded. </font> 
<br /><br />


<strong>Linkage Guide (Redirect Resource Types) </strong>

<br />
The gotoPage contains two required parameters. The first parameter is called URLTYPE. The URLType specifics the target resource type in which to 
navigation within site the site. All possible types are shown below. Based on the URLType, the 2nd Parameter called Value, should be set 
to the resource ID. 
<br /><br />
For example, if linking to the Charbroil Commercial 4 burner Marketing Product page, the link would appear as: <br />
http://www.charbroil.com/gotopage.aspx?<font color=red>URLType=16&Value=463257010</font>
<br /><br />
<strong>Adding items to the Shopping Cart:</strong>
<br />
Additionally, if you'd like to add an item to the shopping cart, you can include the &AddToCart=Y attribute to the query string:<br />
http://www.charbroil.com/gotopage.aspx?URLType=16&Value=463257010<font color=red>&AddToCart=Y</font>
<br /><br />
<strong>Redirect to Pages:</strong><br />
Additionally, if you'd like to redirect to a static page located on the website. The following parameter can be used: 
Note: The URLType parameter must not be passed when redirecting to a specific page.   <br />
<br />
<u>Examples:</u> <br />
Example 1: Relative Path <br />
http://www.charbroil.com/gotopage.aspx?<font color=red>pg=/Infrared/Index.html</font> (Recommended) 
<br /><br />
Example 2: Dynamic Path <br />
http://www.charbroil.com/gotopage.aspx?<font color=red>pg=~/Infrared/Index.html</font> 
<br /><br />
Example 3: Full path Path <br />
http://www.charbroil.com/gotopage.aspx?<font color=red>pg=http://www.charbroil.com/Infrared/Index.html</font>
<br /><br />
Example 4: Outbound links<br />
http://www.charbroil.com/gotopage.aspx?<font color=red>pg=http://www.lowes.com/somepromo/&EC=Retailer&EA=Lowes&EL=463666510&EV=1</font>




<br /><br />

<font color="red">Note: Be sure to test all redirect configuration prior to launch</font>



<table border ="1" cellpadding ="3" cellspacing ="0" >
<tr><td align ="left" ><b>Resource</b> </td><td><b>URLType ID</b></td><td><b>Usage Sample</b></td></tr>
<tr><td>Product Marketing Page by Model Number </td><td>16</td><td>GotoPage.aspx?URLType=16&Value=463257010</td></tr>
<tr><td>Product Ecommerce Page by Model Number </td><td>2</td><td>GotoPage.aspx?URLType=2&Value=463257010</td></tr>
<tr><td>Product Support Page by Model Number </td><td>13</td><td>GotoPage.aspx?URLType=13&Value=463257010</td></tr>
<tr><td>Marketing Series by ID (Charbroil Commercial Series) </td><td>18</td><td>GotoPage.aspx?URLType=18&Value=32</td></tr>
<tr><td>Marketing Category By ID (BBQ's & Grills) </td><td>5</td><td>GotoPage.aspx?URLType=5&Value=1</td></tr>
<tr><td>Marketing Sub Category By ID (Gas Grills) </td><td>6</td><td>GotoPage.aspx?URLType=6&Value=4</td></tr>
<tr><td>Ecommerce Category By ID (Covers)</td><td>21</td><td>GotoPage.aspx?URLType=21&Value=15</td></tr>
<tr><td>Ecommerce Sub Category By ID (Custom Covers) </td><td>22</td><td>GotoPage.aspx?URLType=22&Value=16</td></tr>
<tr><td>Product Parts by Model Number </td><td>4</td><td>GotoPage.aspx?URLType=4&Value=463257010</td></tr>
<tr><td>Product Accessories by Model Number </td><td>9</td><td>GotoPage.aspx?URLType=9&Value=463257010</td></tr>
<tr><td>Article by ID (Cooking Chart) </td><td>11</td><td>GotoPage.aspx?URLType=11&Value=458</td></tr>
<tr><td>Recipe by ID</td><td>14</td><td>GotoPage.aspx?URLType=14&Value=[RecipeID]</td></tr>
<tr><td>Recipe Category By ID </td><td>17</td><td>GotoPage.aspx?URLType=17&Value=[Recipe Category ID]</td></tr>
<tr><td>Specialty Store/Promotion By ID (Great Gift for Dad)</td><td>7</td><td>GotoPage.aspx?URLType=7&Value=21</td></tr>
<tr><td>Store Home Page </td><td>23</td><td>GotoPage.aspx?URLType=23</td></tr>
<tr><td>Clearance Section</td><td>24</td><td>GotoPage.aspx?URLType=24</td></tr>
<tr><td>Retailer links (outbound)</td><td>25</td><td>GotoPage.aspx?URLType=25&Value=[RetailerID]&Value2=[ProductID]</td></tr>
</table>
</div> 
</div> 

</asp:Content>

