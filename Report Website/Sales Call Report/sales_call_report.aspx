<%@ Page Title="Sales Call Report Page" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="sales_call_report.aspx.cs" Inherits="sales_call_report"
    MaintainScrollPositionOnPostback="true" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Sales Call Reports</font>
    </div>
    <div align="center">
        <asp:Button ID="AddButton" runat="server" Font-Bold="True" Font-Size="X-Large" Height="74px"
            Text="Add Report" OnClick="AddButton_Click" Width="150px" />
        <hr />
        <asp:Label ID="CustomerListLabel" runat="server" Font-Bold="True" Text="Sales Call Report For:"></asp:Label>
        <asp:DropDownList ID="CustomerList" runat="server" DataSourceID="cms1" DataTextField="BVNAME"
            DataValueField="BVNAME" Height="29px" Width="218px">
        </asp:DropDownList>
        <asp:SqlDataSource ID="cms1" runat="server" ConnectionString="<%$ ConnectionStrings:cms1 %>"
            ProviderName="<%$ ConnectionStrings:cms1.ProviderName %>" 
            SelectCommand="select distinct trim(bvname) as BVNAME from cmsdat.cust order by trim(bvname)">
        </asp:SqlDataSource>
        <asp:Button ID="CheckButton" runat="server" Font-Bold="True" Font-Size="X-Large"
            Height="74px" Text="Check Report" OnClick="CheckButton_Click" Width="180px" />
        <asp:GridView ID="GridViewReports" runat="server"
            OnSelectedIndexChanged="Select_Click" CellPadding="4" 
            HorizontalAlign="Center" Font-Size="Small"
            AutoGenerateSelectButton="True" ShowHeaderWhenEmpty="True" 
            ForeColor="#333333" Width="656px">
            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
            <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" 
                BackColor="#999999" />
            <FooterStyle BackColor="#5D7B9D" ForeColor="White" Font-Bold="True" />
            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" HorizontalAlign="Center" 
                VerticalAlign="Middle" />
            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
            <SortedAscendingCellStyle BackColor="#E9E7E2" />
            <SortedAscendingHeaderStyle BackColor="#506C8C" />
            <SortedDescendingCellStyle BackColor="#FFFDF8" />
            <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
        </asp:GridView>
    </div>
</asp:Content>
