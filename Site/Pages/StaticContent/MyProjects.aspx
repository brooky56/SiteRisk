<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="MyProjects.aspx.cs" Inherits="Site.Pages.StaticContent.MyProjects" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="Projects" ContentPlaceHolderID="MyProjectsPage" runat="server">
    <div id="wrapper">

        <!-- Sidebar -->
        <div id="sidebar-wrapper">
            <ul class="sidebar-nav">
                <li class="sidebar-brand">
                    <a href="../Projects.aspx">Главная</a>
                </li>
                <li>
                    <a href="MyProjects.aspx">Мои проекты</a>
                </li>
                <li>
                    <a href="NewProject.aspx">Создать проект</a>
                </li>
                <li>
                    <a href="ProjectSettings.aspx">Настройки</a>
                </li>

            </ul>
        </div>
        <!-- /#sidebar-wrapper -->

        <!-- Page Content -->
        <div id="page-content-wrapper">
            <div class="container-fluid">
                <div class="col">
                    <h5>
                        <label>Мои Проекты</label>
                    </h5>
                    <div class="row">
                        <asp:GridView ID="MyProjectsGrid" runat="server" AutoGenerateColumns="false" OnRowCommand="MyProjectsGrid_RowCommand">
                            <Columns>
                                <asp:BoundField DataField="ProjectId" HeaderText="Название проекта" />
                                <asp:BoundField DataField="DataChange" HeaderText="Дата изменений" />
                                <asp:BoundField DataField="TimeWorking" HeaderText="Время работы над проектом" />
                                <asp:BoundField DataField="Status" HeaderText="Статус" />
                                <asp:ButtonField ButtonType="Button" CommandName="Download" HeaderText="Загрузить" Text="Загрузить" ControlStyle-CssClass="btn btn-dark" />
                                <asp:ButtonField ButtonType="Button" CommandName="UpdateProject" HeaderText="Редактировать" Text="Редактировать" ControlStyle-CssClass="btn btn-dark" />
                                <asp:ButtonField ButtonType="Button" CommandName="Copy" HeaderText="Дублировать" Text="Дублировать" ControlStyle-CssClass="btn btn-dark" />
                                <asp:ButtonField ButtonType="Button" CommandName="Delete" HeaderText="Удалить" Text="Удалить" ControlStyle-CssClass="btn btn-danger" />
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
            </div>
        </div>
        <!-- /#page-content-wrapper -->

    </div>
    <!-- /#wrapper -->
    <!-- Menu Toggle Script -->
</asp:Content>
