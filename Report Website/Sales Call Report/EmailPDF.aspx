<%@ Page Title="" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true" CodeFile="EmailPDF.aspx.cs" Inherits="Sales_Call_Report_EmailPDF" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" Runat="Server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" Runat="Server">
    <div align="center">
        <font size="18" color="red">Email Sales Call Report</font>
    </div>
    <div style="height: 50px">       
    </div>
    <div style="width: 735px; height: 119px">        
        <div  align="left" style="height: 59px">
            <asp:Label ID="LblEmail" runat="server" Font-Size="20pt" Height="30px" 
                Text="EmailList " Width="137px" ForeColor="Black"></asp:Label>
                &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <asp:DropDownList ID="DropDownEmailList" runat="server" AutoPostBack="True" 
                Font-Size="20pt" 
                onselectedindexchanged="DropDownEmailList_SelectedIndexChanged">
            </asp:DropDownList>
        </div>
        <div align="left" style="height: 49px">              
            <asp:Button ID="BtnPDF" runat="server" onclick="BtnPDF_Click" 
                Text="Download Attach PDF" />
        </div>
    </div>
    </asp:Content>

