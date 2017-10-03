<%@ Page Language="C#" AutoEventWireup="true" MasterPageFile="../Site.master" CodeFile="monthly_sales_by_customer.aspx.cs"
    Inherits="Monthly_Sales_By_Customer" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Sapa Sales by Month</font>
    </div>
    <br />
    <asp:Button ID="buttonSaveFile" runat="server" Font-Size="X-Large" Height="74px"
        OnClick="ButtonSave_Click" Text="Save File" Width="150px" />
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
            <div align="left">
                <asp:TextBox ID="TextBoxSelectCustomer" runat="server" BorderStyle="None" Font-Bold="True"
                    Font-Size="Large" ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px"
                    Width="319px">Select Customer:</asp:TextBox>
                <asp:RadioButton ID="SelectSapa" Visible="true" Text="Sapa" GroupName="SelectCustomer"
                    runat="server"></asp:RadioButton>
                <asp:RadioButton ID="SelectHydro" Visible="False" Text="Hydro" GroupName="SelectCustomer"
                    runat="server"></asp:RadioButton>
                <br />
            </div>
            <div align="left">
                <asp:TextBox ID="Month" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Select Year/Month:</asp:TextBox>
                <asp:DropDownList ID="YearList" runat="server" />
                <asp:RadioButtonList ID="MonthList" runat="server" RepeatDirection="Horizontal" AutoPostBack="true">
                    <asp:ListItem Selected="True">01</asp:ListItem>
                    <asp:ListItem>02</asp:ListItem>
                    <asp:ListItem>03</asp:ListItem>
                    <asp:ListItem>04</asp:ListItem>
                    <asp:ListItem>05</asp:ListItem>
                    <asp:ListItem>06</asp:ListItem>
                    <asp:ListItem>07</asp:ListItem>
                    <asp:ListItem>08</asp:ListItem>
                    <asp:ListItem>09</asp:ListItem>
                    <asp:ListItem>10</asp:ListItem>
                    <asp:ListItem>11</asp:ListItem>
                    <asp:ListItem>12</asp:ListItem>
                </asp:RadioButtonList>
                <br />
            </div>
            <div align="right">
                <asp:Button ID="GenerateReport" runat="server" Font-Size="X-Large" Height="74px"
                    Text="Generate" OnClick="ButtonGenerate_Click" Width="150px" />
            </div>
            <asp:Panel ID="PanelSalesReport" Visible="true" runat="server">
                <asp:GridView ID="GridViewSalesReport" runat="server" BackColor="White" BorderColor="#3366CC"
                    BorderStyle="None" BorderWidth="1px" CellPadding="4" HorizontalAlign="Center">
                    <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                    <FooterStyle BackColor="#99CCCC" ForeColor="#003399" />
                    <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                    <PagerStyle BackColor="#99CCCC" ForeColor="#003399" HorizontalAlign="Left" />
                    <RowStyle BackColor="White" ForeColor="#003399" HorizontalAlign="Center" VerticalAlign="Middle" />
                    <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                    <SortedAscendingCellStyle BackColor="#EDF6F6" />
                    <SortedAscendingHeaderStyle BackColor="#0D4AC4" />
                    <SortedDescendingCellStyle BackColor="#D6DFDF" />
                    <SortedDescendingHeaderStyle BackColor="#002876" />
                </asp:GridView>
            </asp:Panel>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
