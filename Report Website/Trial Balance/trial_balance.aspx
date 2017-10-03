<%@ Page Title="Trial Balance Page" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="trial_balance.aspx.cs" Inherits="Trial_Balance" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Trial Balance Report</font>
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
                <asp:TextBox ID="TextBoxSelectPlant" runat="server" BorderStyle="None" Font-Bold="True"
                    Font-Size="Large" ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px"
                    Width="137px">Select Plant:</asp:TextBox>
                <asp:RadioButton ID="SelectMarkham" Visible="false" Text="Markham (Fiscal)" GroupName="SelectPlant"
                    runat="server"></asp:RadioButton>
                <asp:RadioButton ID="SelectMichigan" Visible="false" Text="Michigan (Fiscal)" GroupName="SelectPlant"
                    runat="server"></asp:RadioButton>
                <asp:RadioButton ID="SelectTexas" Visible="false" Text="Texas (Fiscal)" GroupName="SelectPlant"
                    runat="server"></asp:RadioButton>
                <asp:RadioButton ID="SelectColombia" Visible="false" Text="Colombia (Calendar)" GroupName="SelectPlant"
                    runat="server"></asp:RadioButton>
                <br />
            </div>
            <div align="left">
                <asp:TextBox ID="Period" runat="server" BorderStyle="None" Font-Bold="True" Font-Size="Large"
                    ForeColor="Blue" Height="56px" Style="text-align: center; margin-top: 0px" Width="137px">Select Period:</asp:TextBox>
                <asp:DropDownList ID="YearList" runat="server">
                </asp:DropDownList>
                <asp:RadioButtonList ID="PeriodList" runat="server" RepeatDirection="Horizontal">
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
                <br />
            </div>
            <div align="right">
                <asp:Button ID="GenerateReport" runat="server" Font-Size="X-Large" Height="74px"
                    Text="Generate" OnClick="ButtonGenerate_Click" Width="150px" />
            </div>
            <asp:Panel ID="PanelTrialBalance" Visible="true" runat="server">
                <asp:GridView ID="GridViewTrialBalance" runat="server" BackColor="White" BorderColor="#3366CC"
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
