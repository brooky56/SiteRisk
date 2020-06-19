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
    public partial class Stage8 : System.Web.UI.Page
    {
        private static DataTable dataTable;
        private static DataTable dataTable2;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                dataTable = new DataTable();

                dataTable.Columns.Add("id");
                dataTable.Columns.Add("Name");
                dataTable.Columns.Add("Description");
                dataTable.Columns.Add("Vulnerability");
                dataTable.Columns.Add("projectId");
                PageProtectMeasureDataBind();

                dataTable2 = new DataTable();
                dataTable2.Columns.Add("id");
                dataTable2.Columns.Add("Name");
                dataTable2.Columns.Add("FullName");
                PageValnerabilityDataBind();

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
                cmd.Parameters.AddWithValue("@status", "Stage8");
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
        protected void PageValnerabilityDataBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = "SELECT * FROM siterisk.vulnerability";
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
                    dr["FullName"] = rdr.GetValue(rdr.GetOrdinal("FullName"));

                    dataTable2.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable2);
                val.DataSource = dv;
                val.DataTextField = "Name";
                val.DataValueField = "FullName";
                val.DataBind();

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

        protected void PageProtectMeasureDataBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.protectivemeasure where projectId={Request.QueryString["projectId"]}";
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
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Description"] = rdr.GetValue(rdr.GetOrdinal("Description"));
                    dr["Vulnerability"] = rdr.GetValue(rdr.GetOrdinal("Vulnerability"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    dataTable.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable);
                protectMeasure.DataSource = dv;
                protectMeasure.DataBind();

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
                string insertQuery = "insert into siterisk.protectivemeasure(Name,Description, Vulnerability,projectId) " +
                    "values(@Name, @Description, @Vulnerability, @projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@Name", defendName.Text);
                cmd.Parameters.AddWithValue("@Description", description.Text);
                cmd.Parameters.AddWithValue("@Vulnerability", val.SelectedItem.Text);
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
            Response.Redirect($"Stage8.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void continueProject_Click(object sender, EventArgs e)
        {
            Response.Redirect($"Stage9.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void protectMeasure_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {

        }

        protected void protectMeasure_RowCommand(object sender, GridViewCommandEventArgs e)
        {

        }
    }
}