<%@ Page Language="VB" Debug ="true"  AutoEventWireup="false" CodeFile="Default.aspx.vb" Inherits="ClientSide_Default" %>

<%@ Register Src="../ClientWebControls/ClientDataCube.ascx" TagName="ClientDataCube"
    TagPrefix="uc1" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
</head>
<body>
    <form id="form1" runat="server">
           <h2>Dynamic DataCube By Category and index</h2>
    <asp:DropDownList runat ="server" ID ="Category" AutoPostBack ="true" >
    <asp:ListItem Value ="0" Text ="Select A Category"></asp:ListItem>
    <asp:ListItem Value ="1" Text ="Category1"></asp:ListItem>
    <asp:ListItem Value ="2" Text ="Category2"></asp:ListItem>
    <asp:ListItem Value ="3" Text ="Category3"></asp:ListItem>
        </asp:DropDownList>
        <asp:Button ID=btnReset runat= "server" Text ="reset" />
        
    <div>
    <table>
        <tr><td colspan =3><uc1:ClientDataCube EnableBorder =true  Width ="500" BorderColor=black BorderSize =1 Index="1" ID="ClientDataCube1" runat="server" /></td></tr>
        <tr>
            <td><uc1:ClientDataCube  Index="2"  ID="ClientDataCube2" runat="server" /></td>
            <td><uc1:ClientDataCube Index="3" ID="ClientDataCube3" runat="server" /></td>
            <td><uc1:ClientDataCube Index="4" ID="ClientDataCube4" runat="server" /></td>
        </tr>
        
    </table>
        
        &nbsp;</div>
        
<br /><br />
       <h2>Static DataCube </h2>
        <uc1:ClientDataCube DataCubeID ="DataCube1" ID="ClientDataCube5" runat="server" />
        <uc1:ClientDataCube DataCubeID ="DataCube2" ID="ClientDataCube6" runat="server" />
    </form>
</body>
</html>
