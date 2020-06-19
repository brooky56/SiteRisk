using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Site.Pages.StaticContent.StepTables
{
    public partial class Stage7 : System.Web.UI.Page
    {
        private static DataTable dataTable;
        private static DataTable dataTable2;
        private static DataTable dataTable3;
        private static DataTable dataTable4;

        protected void Page_Load(object sender, EventArgs e)
        {
            dataTable4 = new DataTable();
            dataTable4.Columns.Add("id");
            dataTable4.Columns.Add("assetType");
            dataTable4.Columns.Add("level0");
            dataTable4.Columns.Add("level1");
            dataTable4.Columns.Add("level2");
            dataTable4.Columns.Add("level3");
            dataTable4.Columns.Add("level4");
            dataTable4.Columns.Add("projectId");

            dataTable = new DataTable();
            dataTable.Columns.Add("assetId");
            dataTable.Columns.Add("assetName");

            dataTable2 = new DataTable();
            dataTable2.Columns.Add("id");
            dataTable2.Columns.Add("Name");
            dataTable2.Columns.Add("Effect");
            dataTable2.Columns.Add("assetType");

            dataTable3 = new DataTable();
            dataTable3.Columns.Add("id");
            dataTable3.Columns.Add("Name");
            dataTable3.Columns.Add("Effect");
            dataTable3.Columns.Add("Price");
            dataTable3.Columns.Add("AssetType");

            PageDataCascadeEventBind();
            PageDataBindWorkEffect();
            PageDataBindEffect();
            PageDataBindAsset();
        }

        protected void PageDataCascadeEventBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.cascadeevent where projectId={Request.QueryString["projectId"]}";
                MySqlCommand cmd = new MySqlCommand(selectquery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };
                con.Open();
                cmd.ExecuteNonQuery();

                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    DataRow dr = dataTable4.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["assetType"] = rdr.GetValue(rdr.GetOrdinal("assetType"));
                    dr["level0"] = rdr.GetValue(rdr.GetOrdinal("level0"));
                    dr["level1"] = rdr.GetValue(rdr.GetOrdinal("level1"));
                    dr["level2"] = rdr.GetValue(rdr.GetOrdinal("level2"));
                    dr["level3"] = rdr.GetValue(rdr.GetOrdinal("level3"));
                    dr["level4"] = rdr.GetValue(rdr.GetOrdinal("level4"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    dataTable4.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable4);
                cascade.DataSource = dv;
                cascade.DataBind();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
            }
        }

        protected void UpdateStageProject()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = $"update siterisk.projects set status=@status where id=@id";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Parameters.AddWithValue("@status", "Stage7");
                cmd.Parameters.AddWithValue("@id", Request.QueryString["projectId"]);
                cmd.ExecuteNonQuery();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
            }
        }

        protected void PageDataBindWorkEffect()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = "SELECT * FROM siterisk.workeffectib";
                MySqlCommand cmd = new MySqlCommand(selectquery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };
                con.Open();
                cmd.ExecuteNonQuery();

                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    DataRow dr = dataTable3.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Effect"] = rdr.GetValue(rdr.GetOrdinal("Effect"));
                    dr["Price"] = rdr.GetValue(rdr.GetOrdinal("Price"));
                    dr["AssetType"] = rdr.GetValue(rdr.GetOrdinal("AssetType"));

                    dataTable3.Rows.Add(dr);
                }

                DataRow dr1 = dataTable3.NewRow();
                dr1["id"] = 0;
                dr1["Name"] = "Нет";
                dataTable3.Rows.Add(dr1);

                DataView dv = new DataView(dataTable3);
                list2.DataSource = dv;
                list2.DataTextField = "Name";
                list2.DataValueField = "id";
                list2.DataBind();

                list3.DataSource = dv;
                list3.DataTextField = "Name";
                list3.DataValueField = "id";
                list3.DataBind();

                list4.DataSource = dv;
                list4.DataTextField = "Name";
                list4.DataValueField = "id";
                list4.DataBind();



            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
            }
        }

        protected void PageDataBindEffect()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = "SELECT * FROM siterisk.threateffectib";
                MySqlCommand cmd = new MySqlCommand(selectquery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };
                con.Open();
                cmd.ExecuteNonQuery();

                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    DataRow dr = dataTable2.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Effect"] = rdr.GetValue(rdr.GetOrdinal("Effect"));
                    dr["assetType"] = rdr.GetValue(rdr.GetOrdinal("assetType"));

                    dataTable2.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable2);
                list1.DataSource = dv;
                list1.DataTextField = "Name";
                list1.DataValueField = "id";
                list1.DataBind();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
            }
        }

        protected void PageDataBindAsset()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = "SELECT * FROM siterisk.assettype";
                MySqlCommand cmd = new MySqlCommand(selectquery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };
                con.Open();
                cmd.ExecuteNonQuery();

                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    DataRow dr = dataTable.NewRow();
                    dr["assetId"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["assetName"] = rdr.GetValue(rdr.GetOrdinal("Name"));

                    dataTable.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable);
                objectList.DataSource = dv;
                objectList.DataTextField = "assetName";
                objectList.DataValueField = "assetId";
                objectList.DataBind();

                list0.DataSource = dv;
                list0.DataTextField = "assetName";
                list0.DataValueField = "assetId";
                list0.DataBind();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
            }
        }

        protected void saveProject_Click(object sender, EventArgs e)
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = "insert into siterisk.cascadeevent(assetType,level0, level1,level2, level3, level4, projectId) " +
                    "values(@assetType, @level0, @level1, @level2, @level3, @level4, @projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@assetType", objectList.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@level0", list0.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@level1", list1.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@level2", list2.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@level3", list3.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@level4", list4.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@projectId", Convert.ToInt32(Request.QueryString["projectId"]));
                cmd.ExecuteScalar();

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
            }
            UpdateStageProject();

            Response.Redirect($"Stage7.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void continueProject_Click(object sender, EventArgs e)
        {
            Response.Redirect($"Stage8.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void cascade_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {

        }

        protected void cascade_RowCommand(object sender, GridViewCommandEventArgs e)
        {

        }
    }
}