<%@ Page Language="VB" AutoEventWireup="false"  ValidateRequest ="false" debug="true"  CodeFile="EditDataCubePage.aspx.vb" Inherits="DataCube_EditDataCubePage"  %>

<%@ Register Src="../WebControls/FormEditDataCube.ascx" TagName="FormEditDataCube"
    TagPrefix="uc1" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/tr/html4/strict.dtd"> 
<HTML>
	<HEAD runat="server" >
		<title>EditDataCubePage</title>
		<meta content="Microsoft Visual Studio .NET 7.1" name="GENERATOR">
		<meta content="Visual Basic .NET 7.1" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<link rel="stylesheet" href="include/default.css" type="text/css">
		<script language="javascript">
<!--
var blnSubmit = true;
function Testlink(url)
{
	window.navigate (url)
	
}

function ShowItems(control,iheight, iwidth)
{
control = "FormEditDataCube1_" + control
var sFeatures = "height=400,width=400,status=no,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes"
	var opener = window.open ("Items/ItemBrowser.aspx?ctr="+control,null,sFeatures)
	opener.focus();
	
}
	
function MergeField(control, typeid)
{

 //Append Form name to control in following format formname_elementid
 control = "FormEditDataCube1_" + control
 
var sFeatures = "height=500,width=250,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes"
	var opener = window.open ("mSelectMergeFields.aspx?ctr="+control+"&tid=" + typeid,"MergeWindow",sFeatures)
	opener.focus();
	
}



function ShowElements(dc,typeid,  control,iheight, iwidth)
{
	
/*	var sFeatures = "dialogHeight: "+ iheight + "px;dialogWidth: " + iwidth + "px;"
	var oParameters = new Object()
	control = "FormEditDataCube1_" + control
	oParameters.callerid = control
	oParameters.datacubeid = dc
	var oElement = window.showModalDialog ("mSelectElement.aspx?ctr="+control + "&did=" + dc + "&tid=" + typeid,oParameters,sFeatures)
*/	
control = "FormEditDataCube1_" + control
var sFeatures = "height=500,width=250,status=yes,toolbar=no,menubar=no,location=no,scrollbars=yes,resizable=yes"
	var opener = window.open ("mSelectElement.aspx?ctr="+control+"&did=" + dc + "&tid=" + typeid,"MergeWindow",sFeatures)
	opener.focus();
	
	 
}

function onSubmit()
{
	if (blnSubmit)
		Form1.submit()
	else
	{
		blnSubmit = true
		return false;
	}
}
//-->
		</script>
	</HEAD>
	<body MS_POSITIONING="FlowLayout">
				<form id="Form1" onsubmit="return onSubmit();" method="post" runat="server">
                    <uc1:FormEditDataCube ID="FormEditDataCube1" runat="server" />
				</form>

</body>
</html>


