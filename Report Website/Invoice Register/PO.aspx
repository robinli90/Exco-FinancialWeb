<%@ Page Title="Vendor Parts Listing" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="PO.aspx.cs" Inherits="invoice_register" %>

<asp:Content ID="Content1" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="Content2" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Vendor Parts Listing</font>
    </div>
    <asp:Button ID="ButtonSave" runat="server" Text="Save" Font-Size="X-Large" Height="74px"
        Visible="true" OnClick="ButtonSave_Click" Width="200px" />
    <br />
    <asp:ScriptManager ID="ScriptManager1" runat="server" AsyncPostBackTimeout="6000" />
    <div align="center">
        <asp:UpdateProgress ID="UpdateProgress1" runat="server" AssociatedUpdatePanelID="UpdatePanel">
            <ProgressTemplate>
                <font size="8" color="red">Please Wait...</font>
                <img alt="" src="../Images/Loading.gif" />
            </ProgressTemplate>
        </asp:UpdateProgress>
    </div>
    <asp:UpdatePanel ID="UpdatePanel" runat="server">
        <ContentTemplate>
              <div align="center">

                 <asp:TextBox ID="Plant1" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Select Plant:</asp:TextBox>
                <asp:RadioButtonList ID="PlantList1" runat="server" RepeatDirection="Horizontal" AutoPostBack="true">
                    <asp:ListItem Selected="True">Markham</asp:ListItem>
                    <asp:ListItem>Michigan</asp:ListItem>
                    <asp:ListItem>Colombia</asp:ListItem>
                    <asp:ListItem>Texas</asp:ListItem>
                </asp:RadioButtonList>
            </br>
            
            <br/>
                <asp:TextBox ID="text" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Part Description Search (optional):</asp:TextBox>

                   
            <br/>
                <asp:TextBox ID="desc_box1" runat="server" ForeColor="Blue" Height="18px" Style="text-align: center; margin-top: 0px" Width="140px">
                </asp:TextBox>
            </br>
            
            <br/>
                <asp:TextBox ID="Year1" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Select Year (Fiscal):</asp:TextBox>
                <asp:RadioButtonList ID="YearChoice1" runat="server" RepeatDirection="Horizontal" AutoPostBack="true">
                    <asp:ListItem Selected="True">All</asp:ListItem>
                    <asp:ListItem>2010</asp:ListItem>
                    <asp:ListItem>2011</asp:ListItem>
                    <asp:ListItem>2012</asp:ListItem>
                    <asp:ListItem>2013</asp:ListItem>
                    <asp:ListItem>2014</asp:ListItem>
                    <asp:ListItem>2015</asp:ListItem>
                    <asp:ListItem>2016</asp:ListItem>
                    <asp:ListItem>2017</asp:ListItem>
                </asp:RadioButtonList>
                <br />
            </br>
                <asp:TextBox ID="TextBox22" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Sort By:</asp:TextBox>
                <asp:RadioButtonList ID="sortby" runat="server" RepeatDirection="Horizontal" AutoPostBack="true">
                    <asp:ListItem Selected="True">Date</asp:ListItem>
                    <asp:ListItem>Vendor</asp:ListItem>
                    <asp:ListItem>Part</asp:ListItem>
                    <asp:ListItem>Description</asp:ListItem>
                    <asp:ListItem>Price</asp:ListItem>
                </asp:RadioButtonList>
                <br />
            
            </br>
            
            <br/>
            </div>
            <div align="center">
                <asp:Button ID="Button1" runat="server" Font-Size="X-Large" Height="74px"
                    Text="Generate" OnClick="ButtonGenerate_Click" Width="150px" />
            </div>
            <br />
            <asp:Panel ID="PanelTrialBalance" Visible="true" runat="server">
                <asp:GridView ID="ReportGrid" runat="server" BackColor="White" BorderColor="#3366CC"
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


<%-- 
<%@ Page Title="Customer Die Information" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="Copy of invoice_register.aspx.cs" Inherits="invoice_register" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Customer Die Purchases (Under Construction)</font>
    </div>
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

                 <asp:TextBox ID="Plant1" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Select Plant:</asp:TextBox>
                <asp:RadioButtonList ID="PlantList1" runat="server" RepeatDirection="Horizontal" AutoPostBack="true">
                    <asp:ListItem Selected="True">Markham</asp:ListItem>
                    <asp:ListItem>Michigan</asp:ListItem>
                    <asp:ListItem>Colombia</asp:ListItem>
                    <asp:ListItem>Texas</asp:ListItem>
                </asp:RadioButtonList>
                <asp:TextBox ID="Size1" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Select Die Size:</asp:TextBox>
                <asp:RadioButtonList ID="SizeList1" runat="server" RepeatDirection="Horizontal" AutoPostBack="true">
                    <%-- <asp:ListItem Selected="True">ALL</asp:ListItem> 
                    <asp:ListItem Selected="True">02</asp:ListItem>
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
                    <asp:ListItem>13</asp:ListItem>
                    <asp:ListItem>14</asp:ListItem>
                    <asp:ListItem>15</asp:ListItem>
                    <asp:ListItem>16</asp:ListItem>
                    <asp:ListItem>17</asp:ListItem>
                    <asp:ListItem>18</asp:ListItem>
                    <asp:ListItem>19</asp:ListItem>
                    <asp:ListItem>20</asp:ListItem>
                    <asp:ListItem>21</asp:ListItem>
                    <asp:ListItem>22</asp:ListItem>
                    <asp:ListItem>23</asp:ListItem>
                    <asp:ListItem>24</asp:ListItem>
                    <asp:ListItem>25</asp:ListItem>
                    <asp:ListItem>26</asp:ListItem>
                </asp:RadioButtonList>
                <asp:TextBox ID="Year1" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Select Year:</asp:TextBox>
                <asp:RadioButtonList ID="YearChoice1" runat="server" RepeatDirection="Horizontal" AutoPostBack="true">
                    <asp:ListItem Selected="True">All</asp:ListItem>
                    <asp:ListItem>2010</asp:ListItem>
                    <asp:ListItem>2011</asp:ListItem>
                    <asp:ListItem>2012</asp:ListItem>
                    <asp:ListItem>2013</asp:ListItem>
                    <asp:ListItem>2014</asp:ListItem>
                    <asp:ListItem>2015</asp:ListItem>
                </asp:RadioButtonList>
                <br />
            </div>

            <div align="center">
                <asp:Button ID="GenerateReport" runat="server" Font-Size="X-Large" 
                        Height="74px" OnClick="ButtonGenerate_Click" Text="Generate" Width="150px" />
            </div>
            <br />

            <asp:Panel ID="DieReport" Visible="true" runat="server">
                <p>
                    <p>
                    </p>
                    <asp:GridView ID="ReportGrid" runat="server" BackColor="White" BorderColor="#3366CC"
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
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
 

 NEXT 
            <asp:Panel ID="PanelInvoiceRegisterPESO" Visible="false" runat="server">
                <p>
                    <h2>
                        <asp:Label ID="ReportTitlePESO" runat="server"></asp:Label>
                    </h2>
                    <p>
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

<%@ Page Title="Invoice Register Page" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="Copy of invoice_register.aspx.cs" Inherits="invoice_register" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Customer Die Purchases (Under Construction)</font>
    </div>
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
             <div align="center">
                <asp:TextBox ID="Month" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="325px">Select Die Size:</asp:TextBox>
            <asp:RadioButtonList ID="SizeList" runat="server" RepeatDirection="Horizontal" AutoPostBack="true">
                    <asp:ListItem Selected="True">02</asp:ListItem>
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
                    <asp:ListItem>13</asp:ListItem>
                    <asp:ListItem>14</asp:ListItem>
                    <asp:ListItem>15</asp:ListItem>
                    <asp:ListItem>16</asp:ListItem>
                    <asp:ListItem>17</asp:ListItem>
                    <asp:ListItem>18</asp:ListItem>
                    <asp:ListItem>19</asp:ListItem>
                    <asp:ListItem>20</asp:ListItem>
                    <asp:ListItem>21</asp:ListItem>
                    <asp:ListItem>22</asp:ListItem>
                    <asp:ListItem>23</asp:ListItem>
                    <asp:ListItem>24</asp:ListItem>
                    <asp:ListItem>25</asp:ListItem>
                    <asp:ListItem>26</asp:ListItem>
                </asp:RadioButtonList>
            </div>
            <asp:Panel ID="PanelInvoiceRegisterCAD" Visible="false" runat="server">
                <p>
                    <h2>
                        <asp:Label ID="ReportTitleCAD" runat="server"></asp:Label>
                    </h2>
                    <p>
                    </p>
                    <p>
                        Legend
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
                        Lengend
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
                        Lengend
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

--%>