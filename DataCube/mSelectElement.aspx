<%@ Page Language="vb" AutoEventWireup="false" CodeFile="mSelectElement.aspx.vb"   Inherits="DataCube_mSelectElement" %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" >
<html>
<head runat="server">
    <title>Select Element</title>
    <meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
    <meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
    <meta name="vs_defaultClientScript" content="JavaScript">
    <meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">

    <script type="text/javascript" language="javascript">
<!--
function UpdateMasterForm()
{
	var opener = window.opener;
	if (Form1.lstElement.options[Form1.lstElement.selectedIndex].value != "")
	{
		try
		{
			var control = eval("opener.document.forms[0]." + Form1.hdnControlName.value)
			control.value = Form1.lstElement.options[Form1.lstElement.selectedIndex].value
			window.close();
		}
		catch(e)
		{
			alert("This option is not enabled for this control type [" + control.tagName + "!")
		}
	}
}


function Close()
{

	var oElement = new Object()
	oElement.elementid = Form1.lstElement.options[Form1.lstElement.selectedIndex].value
	window.returnValue = oElement.elementid
	window.close();
}

function ExitWindow()
{
	var oElement = null
	window.returnValue = oElement
	window.close();
	
}
//-->
    </script>

</head>
<body>
    <form id="Form1" method="post" target="_self" runat="server">
        <table cellpadding="0" cellspacing="0" border="0">
            <tr>
                <td colspan="3">
                    Select an Element from the list below.
                </td>
            </tr>
            <tr>
                <td colspan="3">
                    <asp:DropDownList ID="lstElement" runat="server">
                    </asp:DropDownList></td>
            </tr>
            <tr>
                <td align="center" colspan="3">
                    <input id="btnCancel" onclick="ExitWindow();" type="button" value="Close"><input
                        onclick="UpdateMasterForm();" type="button" id="btnSave" value="Select"></td>
            </tr>
        </table>
        <input type="hidden" id="hdnColumn" runat="server">
        <input type="hidden" id="hdnControlName" runat="server">
    </form>
</body>
</html>
