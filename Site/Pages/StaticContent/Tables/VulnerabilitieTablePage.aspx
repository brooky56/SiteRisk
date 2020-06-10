<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="VulnerabilitieTablePage.aspx.cs" Inherits="Site.Pages.StaticContent.Tables.HazardsTablePage" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="VulTable" ContentPlaceHolderID="VulTablePage" runat="server">
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
                    <a href="HazardsTablePage.aspx">Таблица Угроз</a>
                </li>
                <li>
                    <a href="VulnerabilitieTablePage.aspx">Таблица Уязвимостей</a>
                </li>
                <li>
                    <a href="HazardToValnerabilitiesTablePage.aspx">Таблица Угроз и Уязвимостей</a>
                </li>
                <li>
                    <a href="HazardsToAssetsTablePage.aspx">Таблица Угроз и Активов</a>
                </li>
                <li>
                    <a href="AssetTypesTablePage.aspx">Таблица Активов</a>
                </li>
            </ul>
        </div>
        <!-- /#sidebar-wrapper -->

        <!-- Page Content -->
        <div id="page-content-wrapper">
            <div class="container-fluid">
                <div class="col">
                    <h5>
                        <label>Таблица уязвимостей безопасности информации</label>
                    </h5>
                    <div class="row">
                        <asp:GridView ID="vulGrid" runat="server" AutoGenerateColumns="false"
                            AllowPaging="true" PageSize="5" OnPageIndexChanging="vulGrid_PageIndexChanging"
                            CssClass="mydatagrid"
                            AlternatingRowStyle-CssClass="alt"
                            PagerStyle-CssClass="pgr">

                            <Columns>
                                <asp:BoundField DataField="id" HeaderText="Id" />
                                <asp:BoundField DataField="Name" HeaderText="Условное обозначение" />
                                <asp:BoundField DataField="FullName" HeaderText="Наименование уязвимого звена" />

                                <asp:ButtonField ButtonType="Button" CommandName="UpdateProject" HeaderText="Редактировать" Text="Редактировать" ControlStyle-CssClass="btn btn-dark" />
                                <asp:ButtonField ButtonType="Button" CommandName="Delete" HeaderText="Удалить" Text="Удалить" ControlStyle-CssClass="btn btn-danger" />
                            </Columns>
                        </asp:GridView>
                        <asp:TextBox placeholder="Добавить уязвимость (у.о.)" ID="tb1" runat="server" CssClass="form-control" />
                        <asp:TextBox placeholder="Добавить уязвимость (п.о.)" ID="tb2" runat="server" CssClass="form-control" />
                        <asp:Button CssClass="btn btn-info" Text="Добавить уязвимость" ID="addHazardType" OnClick="addHazardType_Click" runat="server" />
                    </div>
                </div>
            </div>
            <!-- /#page-content-wrapper -->
        </div>
    </div>
    <!-- /#wrapper -->
    <!-- Menu Toggle Script -->
</asp:Content>
