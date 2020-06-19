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
    public partial class Stage6 : System.Web.UI.Page
    {
        private static DataTable dataTable;
        private static DataTable dataTable2;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                dataTable2 = new DataTable();
                dataTable2.Columns.Add("id");
                dataTable2.Columns.Add("Name");
                dataTable2.Columns.Add("Effect");
                dataTable2.Columns.Add("Price");
                dataTable2.Columns.Add("AssetType");
                dataTable2.Columns.Add("level");
                dataTable2.Columns.Add("projectId");
                PageDataEffectWorkBind();

                dataTable = new DataTable();
                dataTable.Columns.Add("assetId");
                dataTable.Columns.Add("assetName");
                PageDataBind();
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
                cmd.Parameters.AddWithValue("@status", "Stage6");
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

        protected void PageDataEffectWorkBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.workeffectib where projectId = {Request.QueryString["projectId"]}";
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
                    dr["Price"] = rdr.GetValue(rdr.GetOrdinal("Price"));
                    dr["AssetType"] = rdr.GetValue(rdr.GetOrdinal("AssetType"));
                    dr["level"] = rdr.GetValue(rdr.GetOrdinal("level"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    dataTable2.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable2);
                workEffect.DataSource = dv;
                workEffect.DataBind();

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
        protected void PageDataBind()
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
                IAType.DataSource = dv;
                IAType.DataTextField = "assetName";
                IAType.DataValueField = "assetId";
                IAType.DataBind();

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
                string insertQuery = "insert into siterisk.workeffectib(Name,Effect, Price, AssetType, level, projectId) " +
                    "values(@Name,@Effect, @Price, @AssetType, @level, @projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@Name", IAName.Text);
                cmd.Parameters.AddWithValue("@Effect", IADescription.Text);
                cmd.Parameters.AddWithValue("@Price", Convert.ToDouble(damage.Text));
                cmd.Parameters.AddWithValue("@AssetType", IAType.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@level", Convert.ToInt32(consLevel.Text));
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
            Response.Redirect($"Stage6.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void continueProject_Click(object sender, EventArgs e)
        {
            Response.Redirect($"Stage7.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void workEffect_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {

        }

        protected void workEffect_RowCommand(object sender, GridViewCommandEventArgs e)
        {

        }
    }
}