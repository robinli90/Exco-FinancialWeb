<%@ Page Title="Submit Sales Call Report Page" Language="C#" MasterPageFile="../Site.master"
    AutoEventWireup="true" CodeFile="submit_sales_call_report.aspx.cs" Inherits="submit_sales_call_report"
    MaintainScrollPositionOnPostback="true" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <div align="center">
        <font size="18" color="red">Sales Call Report</font>
    </div>
    <div style="width: 700px; height: 1477px">
        <asp:Label ID="CustomerListLabel" runat="server" Font-Bold="True" Text="Visitation Report For:"></asp:Label>
        <asp:TextBox ID="CustomerText" runat="server" Width="221px" />
        <asp:DropDownList ID="CustomerList" runat="server" DataSourceID="cms1" DataTextField="BVNAME"
            DataValueField="BVNAME" Height="29px" Width="218px">
        </asp:DropDownList>
        <asp:SqlDataSource ID="cms1" runat="server" ConnectionString="<%$ ConnectionStrings:cms1 %>"
            ProviderName="<%$ ConnectionStrings:cms1.ProviderName %>" SelectCommand="select distinct trim(bvname) as BVNAME from cmsdat.cust order by trim(bvname)">
        </asp:SqlDataSource>
        <hr />
        <div style="width: 655px">
            <asp:Label ID="SalesLabel" runat="server" Font-Bold="True" Text="Sales:"></asp:Label>
            <asp:TextBox ID="SalesText" runat="server"></asp:TextBox>
        </div>
        <hr />
        <div style="width: 655px">
            <asp:Label ID="DateLabel" runat="server" Font-Bold="True" Text="Date:"></asp:Label>
            <asp:TextBox ID="DateText" runat="server"></asp:TextBox>
            <asp:Calendar ID="DateCalendar" runat="server" Width="220px" BackColor="White" BorderColor="#3366CC"
                BorderWidth="1px" CellPadding="1" DayNameFormat="Shortest" Font-Names="Verdana"
                Font-Size="8pt" ForeColor="#003399" Height="200px">
                <DayHeaderStyle BackColor="#99CCCC" ForeColor="#336666" Height="1px" />
                <NextPrevStyle Font-Size="8pt" ForeColor="#CCCCFF" />
                <OtherMonthDayStyle ForeColor="#999999" />
                <SelectedDayStyle BackColor="#009999" Font-Bold="True" ForeColor="#CCFF99" />
                <SelectorStyle BackColor="#99CCCC" ForeColor="#336666" />
                <TitleStyle BackColor="#003399" BorderColor="#3366CC" BorderWidth="1px" Font-Bold="True"
                    Font-Size="10pt" ForeColor="#CCCCFF" Height="25px" />
                <TodayDayStyle BackColor="#99CCCC" ForeColor="White" />
                <WeekendDayStyle BackColor="#CCCCFF" />
            </asp:Calendar>
        </div>
        <hr />
        <div id="attendeediv" style="width: 655px; height: 136px;">
            <asp:Label ID="AttendeeLabel" runat="server" Font-Bold="True" Text="Attendee's:"></asp:Label>
            
            &nbsp;&nbsp;&nbsp;&nbsp;
            <br />
            <asp:TextBox ID="AttendeeText" runat="server" Height="111px" 
                style="margin-left: 0px; margin-top: 0px" Width="646px" 
                TextMode="MultiLine" ViewStateMode="Enabled"></asp:TextBox>
        </div>
        <hr />
        <div style="width: 655px">
            <asp:Label ID="PressLabel" runat="server" Font-Bold="True" Text="Number of Presses:"></asp:Label>
            <asp:TextBox ID="PressText" runat="server"></asp:TextBox>
        </div>
        <hr />
        <div style="width: 655px">
            <asp:Label ID="ExpenseLabel" runat="server" Font-Bold="True" Text="Estimated Die Expense/Mn:"></asp:Label>
            <asp:TextBox ID="ExpenseText" runat="server"></asp:TextBox>
        </div>
        <hr />
        <div style="width: 655px">
            <asp:Label ID="SupplierLabel" runat="server" Font-Bold="True" Text="Other Suppliers:"></asp:Label>
            <asp:TextBox ID="SupplierText" runat="server" Width="427px"></asp:TextBox>
        </div>
        <hr />
        <div style="width: 655px">
            <asp:Label ID="PurposeLabel" runat="server" Font-Bold="True" Text="Purpose of Visit:"></asp:Label>
            <asp:TextBox ID="PurposeText" runat="server" TextMode="MultiLine" 
                Height="135px" Width="646px"></asp:TextBox>
        </div>
        <hr />
        <div style="width: 655px; height: 164px;">
            <asp:Label ID="InformationLabel" runat="server" Font-Bold="True" Text="Addtional Information:"></asp:Label>
            <asp:TextBox ID="InformationText" runat="server" TextMode="MultiLine" Height="134px"
                Width="647px"></asp:TextBox>
        </div>
        <hr />
        <div style="width: 655px; height: 179px;">
            <asp:Label ID="IncreaseLabel" runat="server" Font-Bold="True" Text="How can ETS Increase Business?"></asp:Label>
            <asp:TextBox ID="IncreaseText" runat="server" TextMode="MultiLine" Height="154px"
                Width="645px"></asp:TextBox>
        </div>
        <hr />
        <div style="width: 655px">
            <asp:Label ID="FollowUpLabel" runat="server" Font-Bold="True" Text="Follow Up Date:"></asp:Label>
            <asp:TextBox ID="FollowUpText" runat="server"></asp:TextBox>
            <asp:Calendar ID="FollowUpCalendar" runat="server" Width="220px" BackColor="#FFFFCC"
                BorderColor="#FFCC66" BorderWidth="1px" DayNameFormat="Shortest" Font-Names="Verdana"
                Font-Size="8pt" ForeColor="#663399" Height="200px" ShowGridLines="True">
                <DayHeaderStyle BackColor="#FFCC66" Font-Bold="True" Height="1px" />
                <NextPrevStyle Font-Size="9pt" ForeColor="#FFFFCC" />
                <OtherMonthDayStyle ForeColor="#CC9966" />
                <SelectedDayStyle BackColor="#CCCCFF" Font-Bold="True" />
                <SelectorStyle BackColor="#FFCC66" />
                <TitleStyle BackColor="#990000" Font-Bold="True" Font-Size="9pt" ForeColor="#FFFFCC" />
                <TodayDayStyle BackColor="#FFCC66" ForeColor="White" />
            </asp:Calendar>
        </div>
        <div align="center">
            <asp:Button ID="SubmitButton" runat="server" Font-Bold="True" Text="Submit" OnClick="SubmitButton_Click" />
            <asp:Button ID="ClearButton" runat="server" Font-Bold="True" Text="Clear" OnClick="ClearButton_Click" />
        </div>
        <br />
    </div>
</asp:Content>
