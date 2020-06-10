<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Stage7.aspx.cs" Inherits="Site.Pages.StaticContent.StepTables.Stage7" MasterPageFile="~/Pages/Site.Master"%>

<asp:Content ID="Stage7Table" ContentPlaceHolderID="Stage7Page" runat="server">
    <style>
        .mydatagrid {
            width: 80%;
            border: solid 2px black;
            min-width: 80%;
        }

        .header {
            background-color: #000;
            font-family: Arial;
            color: White;
            height: 25px;
            text-align: center;
            font-size: 16px;
        }

        .rows {
            background-color: #fff;
            font-family: Arial;
            font-size: 14px;
            color: #000;
            min-height: 25px;
            text-align: left;
        }

            .rows:hover {
                background-color: #5badff;
                color: #fff;
            }

        .mydatagrid a /** FOR THE PAGING ICONS **/ {
            background-color: Transparent;
            padding: 5px 5px 5px 5px;
            color: #fff;
            text-decoration: none;
            font-weight: bold;
        }

            .mydatagrid a:hover /** FOR THE PAGING ICONS HOVER STYLES**/ {
                background-color: #000;
                color: #fff;
            }

        .mydatagrid span /** FOR THE PAGING ICONS CURRENT PAGE INDICATOR **/ {
            background-color: #fff;
            color: #000;
            padding: 5px 5px 5px 5px;
        }

        .pager {
            background-color: #5badff;
            font-family: Arial;
            color: White;
            height: 30px;
            text-align: left;
        }

        .mydatagrid td {
            padding: 5px;
        }

        .mydatagrid th {
            padding: 5px;
        }
    </style>
    <div id="wrapper">
        <!-- Sidebar -->
        <div id="sidebar-wrapper">
            <ul class="sidebar-nav">
                <li class="sidebar-brand">
                    <a href="../../Projects.aspx">Главная</a>
                </li>
                <li>
                    <a href="../MyProjects.aspx">Мои проекты</a>
                </li>
                <li>
                    <a href="../NewProject.aspx">Создать проект</a>
                </li>
                <li>
                    <a href="../ProjectSettings.aspx">Настройки</a>
                </li>
                <li>
                    <a href="Stage6.aspx">Назад</a>
                </li>
            </ul>
        </div>
        <!-- /#sidebar-wrapper -->

        <!-- Page Content -->
        <div id="page-content-wrapper">
            <div class="container-fluid">
                <div class="card-body">
                    <div class="row">
                        <div class="col">
                            <h5>
                                <label>Сценарии каскадных событий</label>
                            </h5>
                            <label>Объект воздействия</label>
                            <div class="form-group">
                                <asp:DropDownList CssClass="form-control" ID="objectList" runat="server" />
                            </div>
                            <label>Последствие 0 уровня</label>
                            <div class="form-group">
                                 <asp:DropDownList CssClass="form-control" ID="list0" runat="server" />
                            </div>
                            <label>Последствие 1 уровня</label>
                            <div class="form-group">
                                 <asp:DropDownList CssClass="form-control" ID="list1" runat="server" />
                            </div>
                            <label>Последствия 2 уровня</label>
                            <div class="form-group">
                                <asp:DropDownList CssClass="form-control" ID="list2" runat="server" />
                            </div>
                            <label>Последствия 3 уровня</label>
                            <div class="form-group">
                                 <asp:DropDownList CssClass="form-control" ID="list3" runat="server" />
                            </div>
                            <label>Последствия 4 уровня</label>
                            <div class="form-group">
                                 <asp:DropDownList CssClass="form-control" ID="list4" runat="server" />
                            </div>
                            <div class="form-group">
                                <asp:Button ID="saveProject" Text="Сохранить" runat="server" OnClick="saveProject_Click" CssClass="btn btn-success" />
                            </div>
                            <div class="form-group">
                                <asp:Button ID="continueProject" Text="Продолжить" runat="server" OnClick="continueProject_Click" CssClass="btn btn-warning" />
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        <!-- /#page-content-wrapper -->

    </div>
    <!-- /#wrapper -->
    <!-- Menu Toggle Script -->
</asp:Content>
