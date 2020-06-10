<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="NewProject.aspx.cs" Inherits="Site.Pages.StaticContent.NewProject" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="NewProjects" ContentPlaceHolderID="NewProjectsPage" runat="server">
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
                <div class="card-body">
                    <div class="row">
                        <div class="col">
                            <h5>
                                <label>Общие сведения о проекте</label>
                            </h5>

                            <label>Полное наименование организации </label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="accountName" runat="server" placeholder="Полное наименование организации" />
                            </div>
                            <label>Сокращенное наименование организации </label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="shortNameAccountName" runat="server" placeholder="Сокращенное наименование организации" />
                            </div>
                            <label>Юридический Адрес </label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="accountAddress" runat="server" placeholder="Юридический Адрес" />
                            </div>
                            <label>Фактический Адрес </label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="accountActualAdress" runat="server" placeholder="Фактический Адрес" />
                            </div>
                            <label>Контактное лицо </label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="contact" runat="server" placeholder="Контактное лицо" />
                            </div>
                            <label>Телефон </label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="phone" runat="server" placeholder="Телефон" />
                            </div>
                            <label>E-mail </label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="email" runat="server" placeholder="E-mail" />
                            </div>
                            <label>Приемлемый ущерб </label>
                            <div class="form-group">
                                <asp:TextBox CssClass="form-control" ID="normDamage" runat="server" placeholder="Приемлемый ущерб" />
                            </div>
                            <div class="form-group">
                                <asp:Button ID="saveProject" Text="Сохранить" runat="server" OnClick="saveProject_Click" CssClass="btn btn-success"  />
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
