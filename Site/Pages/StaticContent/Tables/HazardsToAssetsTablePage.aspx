<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="HazardsToAssetsTablePage.aspx.cs" Inherits="Site.Pages.StaticContent.Tables.HazardsToAssetsTablePage" MasterPageFile="~/Pages/Site.Master" %>

<asp:Content ID="HazardsToAssets" ContentPlaceHolderID="HazardsToAssetsPage" runat="server">
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
                            <label>Таблица логических связей между угрозами и типами активов</label>
                        </h5>
                        <asp:GridView ID="threatAssetGrid" runat="server" AutoGenerateColumns="false"
                            AllowPaging="true" PageSize="5" OnPageIndexChanging="threatAssetGrid_PageIndexChanging"
                            CssClass="mydatagrid"
                            AlternatingRowStyle-CssClass="alt"
                            PagerStyle-CssClass="pgr">

                            <Columns>
                                <asp:BoundField DataField="Id" HeaderText="Id" />
                                <asp:BoundField DataField="threatName" HeaderText="Наименование угрозы" />
                                <asp:BoundField DataField="threatShortName" HeaderText="Условное обозначение" />
                                <asp:BoundField DataField="assetName" HeaderText="Тип актива" />

                                <asp:ButtonField ButtonType="Button" CommandName="Delete" HeaderText="Разорвать связь" Text="Удалить" ControlStyle-CssClass="btn btn-danger" />
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
                <div class="row">
                    <div class="col m-2">
                        <asp:DropDownList ID="threatList" runat="server" CssClass="custom-select" />
                    </div>
                    <div class="col m-2">
                        <asp:DropDownList ID="assetList" runat="server" CssClass="custom-select" />
                    </div>
                    <div class="col m-2">
                        <asp:Button ID="AddTALink" Text="Добавить связь" runat="server" OnClick="AddTALink_Click" CssClass="btn btn-info" />
                    </div>
                </div>
            </div>
        </div>
        <!-- /#page-content-wrapper -->

    </div>
    <!-- /#wrapper -->
    <!-- Menu Toggle Script -->
</asp:Content>
