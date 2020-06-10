<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="ProjectSettings.aspx.cs" Inherits="Site.Pages.StaticContent.ProjectSettings" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="ProjectSettings" ContentPlaceHolderID="ProjectSettingsPage" runat="server">
    <div id="wrapper">

        <!-- Sidebar -->
        <div id="sidebar-wrapper">
            <ul class="sidebar-nav">
                <li class="sidebar-brand">
                    <a href="../Projects.aspx">Главная</a>
                </li>
                <li>
                    <a href="Tables/HazardsTablePage.aspx">Таблица Угроз</a>
                </li>
                <li>
                    <a href="Tables/VulnerabilitieTablePage.aspx">Таблица Уязвимостей</a>
                </li>
                <li>
                    <a href="Tables/HazardToValnerabilitiesTablePage.aspx">Таблица Угроз и Уязвимостей</a>
                </li>
                <li>
                    <a href="Tables/HazardsToAssetsTablePage.aspx">Таблица Угроз и Активов</a>
                </li>
                <li>
                    <a href="Tables/AssetTypesTablePage.aspx">Таблица Активов</a>
                </li>
            </ul>
        </div>
        <!-- /#sidebar-wrapper -->

        <!-- Page Content -->
        <div id="page-content-wrapper">
            <div class="container-fluid">
            </div>
        </div>
        <!-- /#page-content-wrapper -->

    </div>
    <!-- /#wrapper -->
    <!-- Menu Toggle Script -->
</asp:Content>
