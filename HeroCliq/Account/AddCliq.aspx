<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="AddCliq.aspx.cs" Inherits="HeroCliq.Account.WebForm1" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <asp:Panel ID="Panel1" runat="server" Height="16px">
        <br />
        <asp:Panel ID="Panel2" runat="server">
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<br />
            <asp:DropDownList ID="ddlCliq" runat="server" AutoPostBack="True">
                <asp:ListItem Value="Re">Red</asp:ListItem>
                <asp:ListItem Value="Gr">Gray</asp:ListItem>
                <asp:ListItem Value="Pu">Purple</asp:ListItem>
                <asp:ListItem Value="Br">Brown</asp:ListItem>
                <asp:ListItem Value="Or">Orange</asp:ListItem>
                <asp:ListItem Value="Bl">Black</asp:ListItem>
                <asp:ListItem Value="LB">Light Blue</asp:ListItem>
                <asp:ListItem Value="DB">Dark Blue</asp:ListItem>
                <asp:ListItem Value="LG">Light Green</asp:ListItem>
                <asp:ListItem Value="DG">Dark Green</asp:ListItem>
                <asp:ListItem Value="Wh">White</asp:ListItem>
                <asp:ListItem>KO</asp:ListItem>
            </asp:DropDownList>
            <br />&nbsp;&nbsp;</asp:Panel>
</asp:Panel>
</asp:Content>
