<%@ Page Title="Income Statement Report Page" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="income_statement_report.aspx.cs" Inherits="income_statement_report" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Income Statement Report </font>&nbsp;</div>
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
                <br />
                <br />
                <asp:TextBox ID="TxtBoxSelectPlant" runat="server" BorderStyle="None" 
                    Font-Bold="True" Font-Size="Large" ForeColor="Blue" Height="56px" 
                    Style="text-align: center; margin-top: 0px" Width="137px">Select Plant:</asp:TextBox>
                <asp:RadioButton ID="SelectConsolidate" runat="server" GroupName="SelectPlant" 
                    Text="Consolidate (Fiscal)" Visible="false" />
                <asp:RadioButton ID="SelectMarkham" runat="server" GroupName="SelectPlant" 
                    Text="Markham (Fiscal)" Visible="false" />
                <asp:RadioButton ID="SelectMichigan" runat="server" GroupName="SelectPlant" 
                    Text="Michigan (Fiscal)" Visible="false" />
                <asp:RadioButton ID="SelectTexas" runat="server" GroupName="SelectPlant" 
                    Text="Texas (Fiscal)" Visible="false" />
                <asp:RadioButton ID="SelectColombia" runat="server" GroupName="SelectPlant" 
                    Text="Colombia (Fiscal)" Visible="false" />
                &nbsp;&nbsp;&nbsp;&nbsp;<br>
                <asp:TextBox ID="Period" runat="server" BorderStyle="None" Font-Bold="True" 
                    Font-Size="Large" ForeColor="Blue" Height="56px" 
                    Style="text-align: center; margin-top: 0px; margin-left: 0px" Width="137px">Select Period:</asp:TextBox>
                <asp:DropDownList ID="YearList" runat="server">
                </asp:DropDownList>
                <br />
                <asp:RadioButtonList ID="PeriodList" runat="server" 
                    RepeatDirection="Horizontal">
                    <asp:ListItem>01</asp:ListItem>
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
                <asp:Button ID="GenerateReport" runat="server" Font-Size="X-Large" 
                    Height="74px" OnClick="ButtonGenerate_Click" Text="Generate" Width="150px" />
                <br />
            </div>
            <asp:Panel ID="PanelIncomeStatementReport" Visible="false" runat="server">
                <asp:GridView ID="GridViewIncomeStatementReport" runat="server" 
                    BackColor="White" BorderColor="#3366CC"
                    BorderStyle="None" BorderWidth="1px" CellPadding="4" 
                    HorizontalAlign="Center" Visible="False">
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
