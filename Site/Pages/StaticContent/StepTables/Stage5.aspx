<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Stage5.aspx.cs" Inherits="Site.Pages.StaticContent.StepTables.Stage5" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="Stage5Table" ContentPlaceHolderID="Stage5Page" runat="server">
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
                    <a href="Stage4.aspx">Назад</a>
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
                                <label>Последствия реализации угроз ИБ</label>
                            </h5>

                            <asp:GridView ID="effect" runat="server" AutoGenerateColumns="false"
                                AllowPaging="true" PageSize="5" OnPageIndexChanging="effect_PageIndexChanging"
                                CssClass="mydatagrid"
                                OnRowCommand="effect_RowCommand"
                                AlternatingRowStyle-CssClass="alt"
                                PagerStyle-CssClass="pgr">

                                <Columns>
                                    <asp:BoundField DataField="id" HeaderText="Id" />
                                    <asp:BoundField DataField="Name" HeaderText="Условное обозначение" />
                                    <asp:BoundField DataField="Effect" HeaderText="Последствие" />
                                    <asp:BoundField DataField="assetType" HeaderText="Тип ИА" />
                                    <asp:BoundField DataField="projectId" HeaderText="Номер проекта" />
                                </Columns>
                            </asp:GridView>

                            <label>Условное обозначение</label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="IAName" runat="server" />
                            </div>
                            <label>Последствие</label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="IADescription" runat="server" />
                            </div>
                            <label>Объект воздействия</label>
                            <div class="form-group">
                                <asp:DropDownList CssClass="form-control" ID="IAType" runat="server" />
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
