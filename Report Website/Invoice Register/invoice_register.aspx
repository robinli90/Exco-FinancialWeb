<%@ Page Title="Invoice Register Page" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="invoice_register.aspx.cs" Inherits="invoice_register" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Invoice Register for Current Month</font>
    </div>
    <br />
    <asp:Button ID="ButtonSave" runat="server" Text="Save" Font-Size="X-Large" Height="74px"
        Visible="true" OnClick="ButtonSave_Click" Width="200px" />
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
            <div align="center">
                <asp:TextBox ID="TextBoxSelectPlant" runat="server" BorderStyle="None" Font-Bold="True"
                    Font-Size="Large" ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px"
                    Width="137px">Select Plant:</asp:TextBox>
                <asp:Button ID="ButtonMarkham" runat="server" Font-Size="X-Large" Height="74px" OnClick="ButtonMarkham_Click"
                    Visible="false" Text="Markham" Width="150px" />
                <asp:Button ID="ButtonMichigan" runat="server" Font-Size="X-Large" Height="74px"
                    Visible="false" OnClick="ButtonMichigan_Click" Text="Michigan" Width="150px" />
                <asp:Button ID="ButtonColombia" runat="server" Font-Size="X-Large" Height="74px"
                    Visible="false" OnClick="ButtonColombia_Click" Text="Colombia" Width="150px" />
                <asp:Button ID="ButtonTexas" runat="server" Font-Size="X-Large" Height="74px" Text="Texas"
                    Visible="false" OnClick="ButtonTexas_Click" Width="150px" />
                <br />
            </div>
            <asp:Panel ID="PanelInvoiceRegisterCAD" Visible="false" runat="server">
                <p>
                    <h2>
                        <asp:Label ID="ReportTitleCAD" runat="server"></asp:Label>
                    </h2>
                    <p>
                    </p>
                    <p>
                        Legend #1
                    </p>
                    <p>
                        <span style="background-color: #FA8072">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;Invoices
                        have not been posted yet
                    </p>
                    <asp:GridView ID="GridViewInvoiceRegisterCAD" runat="server" BackColor="White" BorderColor="#3366CC"
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
                </p>
            </asp:Panel>
            <asp:Panel ID="PanelInvoiceRegisterPESO" Visible="false" runat="server">
                <p>
                    <h2>
                        <asp:Label ID="ReportTitlePESO" runat="server"></asp:Label>
                    </h2>
                    <p>
                    </p>
                    <p>
                        Legend #2
                    </p>
                    <p>
                        <span style="background-color: #FA8072">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;Invoices
                        have not been posted yet
                    </p>
                    <asp:GridView ID="GridViewInvoiceRegisterPESO" runat="server" BackColor="White" BorderColor="#3366CC"
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
                    <p>
                    </p>
                </p>
            </asp:Panel>
            <asp:Panel ID="PanelInvoiceRegisterUSD" Visible="false" runat="server">
                <p>
                    <h2>
                        <asp:Label ID="ReportTitleUSD" runat="server"></asp:Label>
                    </h2>
                    <p>
                    </p>
                    <p>
                        Legend #3
                    </p>
                    <p>
                        <span style="background-color: #FA8072">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</span>&nbsp;&nbsp;Invoices
                        have not been posted yet
                    </p>
                    <asp:GridView ID="GridViewInvoiceRegisterUSD" runat="server" BackColor="White" BorderColor="#3366CC"
                        BorderStyle="None" BorderWidth="1px" CellPadding="4" HorizontalAlign="Center">
                        <EditRowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                        <HeaderStyle BackColor="#003399" Font-Bold="True" ForeColor="#CCCCFF" />
                        <RowStyle BackColor="White" ForeColor="#003399" HorizontalAlign="Center" VerticalAlign="Middle" />
                        <SelectedRowStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                    </asp:GridView>
                    <p>
                    </p>
                </p>
            </asp:Panel>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
