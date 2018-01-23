<%@ Page Language="VB" AutoEventWireup="false" CodeFile="ProductSelector.aspx.vb" Inherits="Product_ProductSelector" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Untitled Page</title>
    <script language ="javascript" src ="../Scripts/script.js"></script>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    Browse By Category: <asp:DropDownList ID=ddCategory runat=server ></asp:DropDownList>
    
    <asp:DataList ID="dlProducts" runat="server" >
    
    </asp:DataList>
    
    </div>
    
    </form>
</body>
</html>
