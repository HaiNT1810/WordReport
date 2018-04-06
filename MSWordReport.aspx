<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MSWordReport.aspx.cs" Inherits="MSWordReport.MSWordReport" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title></title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:Button ID="Export" runat="server" OnClick="Export_Click" Text="Export"/>
        <br />
        <asp:Label ID="lbMessage" runat="server"></asp:Label>
        <asp:Label ID="Label" runat="server"></asp:Label>
    </div>
    </form>
</body>
</html>
