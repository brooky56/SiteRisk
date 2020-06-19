<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Stage8.aspx.cs" Inherits="Site.Pages.StaticContent.StepTables.Stage8" MasterPageFile="~/Pages/Site.Master"%>


<asp:Content ID="Stage8Table" ContentPlaceHolderID="Stage8Page" runat="server">
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
                    <a href="Stage7.aspx">Назад</a>
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
                                <label>Защитные меры</label>
                            </h5>
                            

                            <asp:GridView ID="protectMeasure" runat="server" AutoGenerateColumns="false"
                                AllowPaging="true" PageSize="5" OnPageIndexChanging="protectMeasure_PageIndexChanging"
                                CssClass="mydatagrid"
                                OnRowCommand="protectMeasure_RowCommand"
                                AlternatingRowStyle-CssClass="alt"
                                PagerStyle-CssClass="pgr">

                                <Columns>
                                    <asp:BoundField DataField="id" HeaderText="Id" />
                                    <asp:BoundField DataField="Name" HeaderText="Обозначение защитной меры" />
                                    <asp:BoundField DataField="Description" HeaderText="Описание защитной меры" />
                                    <asp:BoundField DataField="Vulnerability" HeaderText="Устраняемая уязвимость" />
                                    <asp:BoundField DataField="projectId" HeaderText="Номер проекта" />
                                </Columns>
                            </asp:GridView>

                            <label>Обозначение защитной меры</label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="defendName" runat="server" />
                            </div>
                            <label>Описание защитной меры</label>
                            <div class="form-group">
                                 <asp:TextBox CssClass="form-control" ID="description" runat="server" />
                            </div>
                            <label>Устраняемая уязвимость</label>
                            <div class="form-group">
                                 <asp:DropDownList CssClass="form-control" ID="val" runat="server" />
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
