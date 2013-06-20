<%@ Page Title="" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="CliqList.aspx.cs" Inherits="HeroCliq.Account.CliqList" %>
<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .style2
        {
            width: 99px;
        }
        .style4
        {
            width: 45px;
        }
    </style>
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

    <p>
        <asp:ListBox ID="lstCliqs" runat="server" AutoPostBack="True" 
            DataSourceID="SqlDataSource1" DataTextField="Cliq_Number" 
            DataValueField="Cliq_Number" 
            onselectedindexchanged="lstCliqs_SelectedIndexChanged" 
            ontextchanged="lstCliqs_SelectedIndexChanged"></asp:ListBox>
        <asp:SqlDataSource ID="SqlDataSource1" runat="server" 
            ConnectionString="<%$ ConnectionStrings:HeroCliq_CliqsConnectionString %>" 
            SelectCommand="SELECT * FROM [Cliqs]"></asp:SqlDataSource>
    </p>
    <asp:Panel ID="Panel1" runat="server">
    </asp:Panel>
    <p>
    </p>
    <p>
        &nbsp;</p>
    
        
    
</asp:Content>
