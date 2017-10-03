<%@ Page Title="Change Password" Language="C#" MasterPageFile="~/Site.master" AutoEventWireup="true"
    CodeFile="ChangePassword.aspx.cs" Inherits="Account_ChangePassword" %>

<asp:Content ID="HeaderContent" runat="server" ContentPlaceHolderID="HeadContent">
</asp:Content>
<asp:Content ID="BodyContent" runat="server" ContentPlaceHolderID="MainContent">
    <h2>
        Change Password
    </h2>
    <p>
        Use the form below to change your password.
    </p>
    <p>
        New passwords are required to be a minimum of 6 characters in length.</p>
    <p>
        <asp:Label ID="OldPasswordLabel" runat="server" Text="Old Password"></asp:Label>
    </p>
    <p>
        <asp:TextBox ID="OldPassword" runat="server" TextMode="Password"></asp:TextBox>
        <asp:Label ID="OldPasswordErrorLabel" runat="server" ForeColor="Red"></asp:Label>
    </p>
    <p>
        <asp:Label ID="NewPasswordLabel" runat="server" Text="New Password"></asp:Label>
    </p>
    <p>
        <asp:TextBox ID="NewPassword" runat="server" TextMode="Password"></asp:TextBox>
        <asp:Label ID="NewPasswordErrorLabel" runat="server" ForeColor="Red"></asp:Label>
    </p>
    <p>
        <asp:Label ID="RepeatPasswordLabel" runat="server" Text="Repeat Password"></asp:Label>
    </p>
    <p>
        <asp:TextBox ID="RepeatPassword" runat="server" TextMode="Password"></asp:TextBox>
    </p>
    <p>
        <asp:Button ID="Cancel" runat="server" Text="Cancel" onclick="Cancel_Click" />
        <asp:Button ID="Update" runat="server" Text="Update" onclick="Update_Click" />
    </p>
    </asp:Content>