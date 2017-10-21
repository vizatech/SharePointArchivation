<%@ Assembly Name="$SharePoint.Project.AssemblyFullName$" %>
<%@ Import Namespace="Microsoft.SharePoint.ApplicationPages" %>
<%@ Register Tagprefix="SharePoint" Namespace="Microsoft.SharePoint.WebControls" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="Utilities" Namespace="Microsoft.SharePoint.Utilities" Assembly="Microsoft.SharePoint, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Register Tagprefix="asp" Namespace="System.Web.UI" Assembly="System.Web.Extensions, Version=4.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35" %>
<%@ Import Namespace="Microsoft.SharePoint" %>
<%@ Assembly Name="Microsoft.Web.CommandUI, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c" %>
<%@ Page Language="C#" AutoEventWireup="true"  CodeBehind="ArchivationViewPage.aspx.cs" Inherits="_1Plus1Archivation.Layouts.BusinessLevel.ArchivationViewPage" DynamicMasterPageFile="~masterurl/default.master" %>

<asp:Content ID="Main" ContentPlaceHolderID="PlaceHolderMain" runat="server">
    <asp:UpdatePanel runat="server">
        <ContentTemplate>

            <asp:Label ID="h2Element" runat="server">Выберите дату (дату завершения договора) старше которой записи будут архивированы</asp:Label>
            <br /><br />
            <asp:Calendar ID="ArchiveItemsDate" runat="server" OnSelectionChanged="ArchiveItemsDate_SelectionChanged"></asp:Calendar>
            
            <h2>Для архивирования отобраны следующие записи:</h2>

            <br /><br />
            <asp:Label ID="ItemIDs" runat="server" ClientIDMode="Static"></asp:Label>
            <br/>
        </ContentTemplate>
    </asp:UpdatePanel>
       
    <asp:Button runat="server" ID="ArchiveItems" Text="Архивировать" ClientIDMode="Static" OnClick="ArchiveItems_Click" />
    <input type="button" onclick="location.href ='http://vizatech.westeurope.cloudapp.azure.com/sites/team';" value="Oтменить"/>
</asp:Content>

