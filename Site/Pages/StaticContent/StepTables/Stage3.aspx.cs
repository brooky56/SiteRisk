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
    public partial class Stage3 : System.Web.UI.Page
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
                dataTable2.Columns.Add("Description");
                dataTable2.Columns.Add("assetType");
                dataTable2.Columns.Add("Price");
                dataTable2.Columns.Add("Time");
                dataTable2.Columns.Add("MainAmount");
                dataTable2.Columns.Add("ReserveAmount");
                dataTable2.Columns.Add("ZipAmount");
                dataTable2.Columns.Add("projectId");
                PageDataBindAsset();

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
                cmd.Parameters.AddWithValue("@status", "Stage3");
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

        protected void continueProject_Click(object sender, EventArgs e)
        {
            
            Response.Redirect($"Stage4.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void saveProject_Click(object sender, EventArgs e)
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = "insert into siterisk.assets(Name,Description,assetType,Price, Time, MainAmount, ReserveAmount, ZipAmount, projectId) " +
                    "values(@Name,@Description,@assetType,@Price, @Time, @MainAmount, @ReserveAmount, @ZipAmount, @projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@Name", IAName.Text);
                cmd.Parameters.AddWithValue("@Description", IADescription.Text);
                cmd.Parameters.AddWithValue("@assetType", IAType.SelectedValue.ToString());
                cmd.Parameters.AddWithValue("@Price", Convert.ToDouble(amount.Text));
                cmd.Parameters.AddWithValue("@Time", Convert.ToInt32(timeDecline.Text));
                cmd.Parameters.AddWithValue("@MainAmount", Convert.ToInt32(mainCount.Text));
                cmd.Parameters.AddWithValue("@ReserveAmount", Convert.ToInt32(addCount.Text));
                cmd.Parameters.AddWithValue("@ZipAmount", Convert.ToInt32(zipCount.Text));
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
            Response.Redirect($"Stage3.aspx?projectId={Request.QueryString["projectId"]}");
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

        protected void PageDataBindAsset()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.assets Where projectId = {Request.QueryString["projectId"]}";
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
                    dr["Description"] = rdr.GetValue(rdr.GetOrdinal("Description"));
                    dr["assetType"] = rdr.GetValue(rdr.GetOrdinal("assetType"));
                    dr["Price"] = rdr.GetValue(rdr.GetOrdinal("Price"));
                    dr["Time"] = rdr.GetValue(rdr.GetOrdinal("Time"));
                    dr["MainAmount"] = rdr.GetValue(rdr.GetOrdinal("MainAmount"));
                    dr["ReserveAmount"] = rdr.GetValue(rdr.GetOrdinal("ReserveAmount"));
                    dr["ZipAmount"] = rdr.GetValue(rdr.GetOrdinal("ZipAmount"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    dataTable2.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable2);
                assetID.DataSource = dv;
                assetID.DataBind();

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

        protected void assetID_RowCommand(object sender, GridViewCommandEventArgs e)
        {

        }

        protected void assetID_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {

        }
    }
}