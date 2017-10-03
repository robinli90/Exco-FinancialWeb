<%@ Page Title="Sales Report Page" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="sales_report.aspx.cs" Inherits="sales_report"
    MaintainScrollPositionOnPostback="true" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Sales Report</font>
    </div>
    <br />
    <asp:Button ID="ButtonSave" runat="server" Text="Save" Font-Size="X-Large" Height="74px"
        Visible="true" OnClick="ButtonSave_Click" Width="200px" />
    <br />
    <asp:TextBox ID="Period" runat="server" BorderStyle="None" Font-Bold="True" 
        Font-Size="Large" ForeColor="Blue" Height="56px" 
        Style="text-align: center; margin-top: 0px; margin-left: 0px" Width="137px">Select Year:</asp:TextBox>
    <asp:DropDownList ID="YearList" runat="server">
    </asp:DropDownList>
    <br />
    <asp:ScriptManager ID="ScriptManager" runat="server" AsyncPostBackTimeout="6000" />
    <div align="center">
        <asp:UpdateProgress ID="UpdateProgress" runat="server" AssociatedUpdatePanelID="UpdatePanel">
            <ProgressTemplate>
                <font size="8" color="red">Please Wait...</font>
                <img alt="" src="../Images/Loading.gif" />
            </ProgressTemplate>
        </asp:UpdateProgress>
    </div>
    <asp:UpdatePanel ID="UpdatePanel" runat="server">
        <ContentTemplate>
            <asp:TextBox ID="TextBoxSelectPlant" runat="server" BorderStyle="None" Font-Bold="True"
                Font-Size="Large" ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px"
                Width="137px">Select Plant:</asp:TextBox>
            <asp:Button ID="ButtonMarkham" runat="server" Font-Size="X-Large" Height="74px" OnClick="Button_Click"
                Text="Markham" Width="150px" />
            <asp:Button ID="ButtonMichigan" runat="server" Font-Size="X-Large" Height="74px"
                OnClick="Button_Click" Text="Michigan" Width="150px" />
            <asp:Button ID="ButtonColombia" runat="server" Font-Size="X-Large" Height="74px"
                OnClick="Button_Click" Text="Colombia" Width="150px" />
            <asp:Button ID="ButtonTexas" runat="server" Font-Size="X-Large" Height="74px" Text="Texas"
                OnClick="Button_Click" Width="150px" />
            <br />
            <asp:GridView ID="GridViewSalesPESO" runat="server" BackColor="White" Visible="false"
                BorderColor="#3366CC" BorderStyle="None" BorderWidth="1px" CellPadding="4" HorizontalAlign="Center"
                Font-Size="Small">
                <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                <RowStyle BackColor="White" ForeColor="#003399" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
            <asp:GridView ID="GridViewSalesCAD" runat="server" BackColor="White" BorderColor="#3366CC"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="Small" HorizontalAlign="Center"
                Visible="false">
                <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                <RowStyle BackColor="White" ForeColor="#003399" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
            <asp:GridView ID="GridViewSalesUSD" runat="server" BackColor="White" BorderColor="#3366CC"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="Small" HorizontalAlign="Center"
                Visible="false">
                <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                <RowStyle BackColor="White" ForeColor="#003399" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
            <asp:GridView ID="GridViewSalesConsolidate" runat="server" BackColor="White" BorderColor="#3366CC"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="Small" HorizontalAlign="Center"
                Visible="false">
                <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                <RowStyle BackColor="White" ForeColor="#003399" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
            <asp:GridView ID="GridViewSurcharge" runat="server" BackColor="White" BorderColor="#3366CC"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="Small" HorizontalAlign="Center"
                Visible="false">
                <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                <RowStyle BackColor="White" ForeColor="#003399" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
            <asp:GridView ID="GridViewPieceCount" runat="server" BackColor="White" BorderColor="#3366CC"
                BorderStyle="None" BorderWidth="1px" CellPadding="4" Font-Size="Small" HorizontalAlign="Center"
                Visible="false">
                <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                <RowStyle BackColor="White" ForeColor="#003399" HorizontalAlign="Center" VerticalAlign="Middle" />
            </asp:GridView>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
