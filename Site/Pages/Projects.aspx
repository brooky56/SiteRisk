<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Projects.aspx.cs" Inherits="Site.Pages.Projects" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="Projects" ContentPlaceHolderID="ProjectsPage" runat="server">
    <div id="wrapper">

        <!-- Sidebar -->
        <div id="sidebar-wrapper">
            <ul class="sidebar-nav">
                <li class="sidebar-brand">
                    <a href="Projects.aspx">Главная</a>
                </li>
                <li>
                    <a href="StaticContent/MyProjects.aspx">Мои проекты</a>
                </li>
                <li>
                    <a href="StaticContent/NewProject.aspx">Создать проект</a>
                </li>
                <li>
                    <a href="StaticContent/ProjectSettings.aspx">Настройки</a>
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
