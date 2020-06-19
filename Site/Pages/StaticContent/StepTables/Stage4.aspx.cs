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


    public partial class Stage4 : System.Web.UI.Page
    {
        private static DataTable dataTable;
        private static DataTable dataTable2;
        private static DataTable dataTable3;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                dataTable = new DataTable();
                dataTable.Columns.Add("threatId");
                dataTable.Columns.Add("threatName");
                dataTable.Columns.Add("threatFullName");

                dataTable2 = new DataTable();
                dataTable2.Columns.Add("assetId");
                dataTable2.Columns.Add("assetName");


                dataTable3 = new DataTable();
                dataTable3.Columns.Add("id");
                dataTable3.Columns.Add("threatName");
                dataTable3.Columns.Add("threatShortName");
                dataTable3.Columns.Add("assetName");
                dataTable3.Columns.Add("projectId");

                PageDataBindThreat();
                PageDataBindAsset();

                PageDataBindAssetThreatTable();

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
                cmd.Parameters.AddWithValue("@status", "Stage4");
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

        protected void PageDataBindAssetThreatTable()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.threatasset where projectId = {Request.QueryString["projectId"]};";
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
                    dr["threatName"] = rdr.GetValue(rdr.GetOrdinal("threatName"));
                    dr["threatShortName"] = rdr.GetValue(rdr.GetOrdinal("threatShortName"));
                    dr["assetName"] = rdr.GetValue(rdr.GetOrdinal("assetName"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));
                    dataTable3.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable3);
                assetThreat.DataSource = dv;
                assetThreat.DataBind();

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
                    DataRow dr = dataTable2.NewRow();
                    dr["assetId"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["assetName"] = rdr.GetValue(rdr.GetOrdinal("Name"));

                    dataTable2.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable2);
                assetList.DataSource = dv;
                assetList.DataTextField = "assetName";
                assetList.DataValueField = "assetId";
                assetList.DataBind();

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

        protected void PageDataBindThreat()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = "SELECT * FROM siterisk.threats";
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
                    dr["threatId"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["threatName"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["threatFullName"] = rdr.GetValue(rdr.GetOrdinal("FullName"));

                    dataTable.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable);
                threatList.DataSource = dv;
                threatList.DataTextField = "threatName";
                threatList.DataValueField = "threatFullName";
                threatList.DataBind();

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

        protected void AddTALink_Click(object sender, EventArgs e)
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = "insert into siterisk.threatasset(threatName,threatShortName,assetName,projectId) " +
                    "values(@threatName,@threatShortName,@assetName, @projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@threatName", threatList.SelectedItem.Value.ToString());
                cmd.Parameters.AddWithValue("@threatShortName", threatList.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@assetName", assetList.SelectedValue.ToString());
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
        }

        protected void saveProject_Click(object sender, EventArgs e)
        {
            UpdateStageProject();
            Response.Redirect($"Stage4.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void continueProject_Click(object sender, EventArgs e)
        {
            
            Response.Redirect($"Stage5.aspx?projectId={Request.QueryString["projectId"]}");
        }
    }
}