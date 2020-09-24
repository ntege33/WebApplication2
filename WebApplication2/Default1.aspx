<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Default1.aspx.cs" Inherits="WebApplication2.Default1" %>

<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">  
<head runat="server">  
    <title>Import/Export/Read CSV file</title>  
</head>  
<body>  
    <form id="form1" runat="server">  
        <div>  
            <div id="dvforgeneratingexcel">  
                <div>  
                    <asp:GridView ID="wdgList" runat="server" Height="410px" AutoGenerateColumns="False"  
                        Style="width: 99.8%;">  
                        <Columns>  
                            <asp:BoundField DataField="ItemType" HeaderText="Item Type"></asp:BoundField> 
                            <asp:BoundField DataField="OrderPriority" HeaderText="Order Priority"></asp:BoundField>
                            <asp:BoundField DataField="OrderDate" HeaderText="Order Date"></asp:BoundField>  
                            <asp:BoundField DataField="UnitsSold" HeaderText="Units Sold"></asp:BoundField>  
                            <asp:BoundField DataField="UnitPrice" HeaderText="Unit Price"></asp:BoundField> 
                            <asp:BoundField DataField="TotalRevenue" HeaderText="Total Cost"></asp:BoundField>  
                            <asp:BoundField DataField="TotalCost" HeaderText="Total Cost"></asp:BoundField>  
                        </Columns>  
  
                    </asp:GridView>  
                </div>  
                <div class="pull-right">  
                    <asp:Button runat="server" ID="btnExport" Text="Export" OnClick="btnExport_OnClick" />  
                   </div>  
            </div>  
            <div id="dvforexcelimport">  
                <asp:FileUpload Width="300" ID="FileUpload1" CssClass="form-control" runat="server" />  
                <asp:Button ID="btn_import" runat="server" CssClass="btn btn-default" Text="Upload Excel sheet" OnClick="btn_import_Click" />  
            </div>  
  
        </div>  
    </form>  
</body>  
</html>  
