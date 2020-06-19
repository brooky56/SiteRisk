<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Stage4.aspx.cs" Inherits="Site.Pages.StaticContent.StepTables.Stage4" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="Stage4Table" ContentPlaceHolderID="Stage4Page" runat="server">
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
                    <a href="Stage3.aspx">Назад</a>
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
                                <label>Установление логических связей между типами активов и угрозами</label>
                            </h5>
                            <div class="col">
                                <div class="row">
                                    <asp:GridView ID="assetThreat" runat="server" AutoGenerateColumns="false"
                                        AllowPaging="true" PageSize="5"
                                        CssClass="mydatagrid"
                                        AlternatingRowStyle-CssClass="alt"
                                        PagerStyle-CssClass="pgr">

                                        <columns>
                                    <asp:BoundField DataField="id" HeaderText="Id" />
                                    <asp:BoundField DataField="threatName" HeaderText="Наименование угрозы" />
                                    <asp:BoundField DataField="threatShortName" HeaderText="Условное обозначение угрозы" />
                                    <asp:BoundField DataField="assetName" HeaderText="Тип актива" />
                                    <asp:BoundField DataField="projectId" HeaderText="Номер проекта" />

                                </columns>
                                    </asp:GridView>

                                    <div class="col m-2">
                                        <label>Угрозы</label>
                                        <asp:DropDownList ID="threatList" runat="server" CssClass="custom-select" />
                                    </div>
                                    <div class="col m-2">
                                        <label>Активы</label>
                                        <asp:DropDownList ID="assetList" runat="server" CssClass="custom-select" />
                                    </div>
                                    <div class="col m-2">
                                        <label>Добавить</label>
                                        <asp:Button ID="AddTALink" Text="Добавить связь" runat="server" OnClick="AddTALink_Click" CssClass="btn btn-info" />
                                    </div>
                                </div>
                            </div>
                            <div class="col">
                                <asp:Button ID="saveProject" Text="Сохранить" runat="server" OnClick="saveProject_Click" CssClass="btn btn-success" />
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
