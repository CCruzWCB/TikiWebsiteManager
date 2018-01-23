<%@ Page Language="VB" AutoEventWireup="false" CodeFile="Calendar.aspx.vb" Inherits="WebControls_Calendar" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/tr/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Calendar</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
    <asp:Literal ID="lblDate" runat="server"></asp:Literal>
    <asp:Calendar ID="Calendar" runat="server" ></asp:Calendar>
        </div>
    </form>
</body>
</html>
