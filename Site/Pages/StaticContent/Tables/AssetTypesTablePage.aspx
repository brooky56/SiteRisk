<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="AssetTypesTablePage.aspx.cs" Inherits="Site.Pages.StaticContent.Tables.AssetsType" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="AssetTypes" ContentPlaceHolderID="AssetTypesPage" runat="server">
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
                <div class="row">
                    <div class="col">
                        <h5>
                            <label>Таблица Типы активов</label>
                        </h5>

                        <asp:GridView ID="assetsGrid" runat="server" AutoGenerateColumns="false"
                            AllowPaging="true" PageSize="5" OnPageIndexChanging="assetsGrid_PageIndexChanging"
                            CssClass="mydatagrid"
                            AlternatingRowStyle-CssClass="alt"
                            PagerStyle-CssClass="pgr">

                            <Columns>
                                <asp:BoundField DataField="assetId" HeaderText="Id" />
                                <asp:BoundField DataField="assetName" HeaderText="Тип актива" />
                            </Columns>

                        </asp:GridView>

                        <asp:TextBox placeholder="Добавить тип" ID="tb1" runat="server" CssClass="form-control" />

                        <asp:Button CssClass="btn btn-info" Text="Добавить тип" ID="addAssetType" OnClick="addAssetType_Click" runat="server" />
                    </div>
                </div>
            </div>
            <!-- /#page-content-wrapper -->
        </div>
    </div>
    <!-- /#wrapper -->
    <!-- Menu Toggle Script -->
</asp:Content>
