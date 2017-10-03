<%@ Page Title="Home Page" Language="C#" MasterPageFile="Site.master" AutoEventWireup="true"
    CodeFile="Default.aspx.cs" Inherits="_Default" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <h2>
        Welcome to EXCO Reporting System!
    </h2>
    <h1>
        This website intends to be accessed within EXCO Technologies Limited internal network.
    </h1>
    <p>
        Contact
        <asp:HyperLink ID="Hyperlink2" Text="IT Department" NavigateUrl="mailto:zwang@etsdies.com"
            runat="server" />. You can change your password
        <asp:HyperLink ID="HyperLinkChangePassword" runat="server">here</asp:HyperLink>.
    </p>
</asp:Content>
