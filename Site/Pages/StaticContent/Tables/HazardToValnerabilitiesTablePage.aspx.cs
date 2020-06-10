using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Site.Pages.StaticContent.Tables
{
    public partial class HazardToValnerabilities : System.Web.UI.Page
    {
        private static DataTable dataTable;
        private static DataTable dataTable2;
        private static DataTable dataTable3;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                dataTable3 = new DataTable();
                dataTable3.Columns.Add("id");
                dataTable3.Columns.Add("Name");
                dataTable3.Columns.Add("FullName");

                dataTable2 = new DataTable();
                dataTable2.Columns.Add("threatId");
                dataTable2.Columns.Add("threatName");
                dataTable2.Columns.Add("threatFullName");


                dataTable = new DataTable();
                dataTable.Columns.Add("id");
                dataTable.Columns.Add("threatName");
                dataTable.Columns.Add("threatShortName");
                dataTable.Columns.Add("vulnerabilityName");
                dataTable.Columns.Add("vulnerabilityShortName");

                PageDataBind();
                PageDataBindThreat();
                PageDataBindVul();
            }
        }

        protected void HTVGrid_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {
            HTVGrid.PageIndex = e.NewPageIndex;

            PageDataBind();

        }

        protected void PageDataBindVul()
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
                    DataRow dr = dataTable3.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["FullName"] = rdr.GetValue(rdr.GetOrdinal("FullName"));

                    dataTable3.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable3);
                vList.DataSource = dv;
                vList.DataTextField = "Name";
                vList.DataValueField = "FullName";
                vList.DataBind();

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
                    DataRow dr = dataTable2.NewRow();
                    dr["threatId"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["threatName"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["threatFullName"] = rdr.GetValue(rdr.GetOrdinal("FullName"));

                    dataTable2.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable2);
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

        protected void PageDataBind()
        {

            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = "SELECT * FROM siterisk.threatvulnerability";
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
                    dr["threatName"] = rdr.GetValue(rdr.GetOrdinal("threatName"));
                    dr["threatShortName"] = rdr.GetValue(rdr.GetOrdinal("threatShortName"));
                    dr["vulnerabilityName"] = rdr.GetValue(rdr.GetOrdinal("vulnerabilityName"));
                    dr["vulnerabilityShortName"] = rdr.GetValue(rdr.GetOrdinal("vulnerabilityShortName"));

                    dataTable.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable);
                HTVGrid.DataSource = dv;
                HTVGrid.DataBind();

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

        protected void AddHVLink_Click(object sender, EventArgs e)
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = "insert into siterisk.threatvulnerability(threatName,threatShortName, vulnerabilityName, vulnerabilityShortName) " +
                    "values(@threatName,@threatShortName, @vulnerabilityName, @vulnerabilityShortName)";
                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@threatName", threatList.SelectedItem.Value);
                cmd.Parameters.AddWithValue("@threatShortName", threatList.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@vulnerabilityName", vList.SelectedItem.Value);
                cmd.Parameters.AddWithValue("@vulnerabilityShortName", vList.SelectedItem.Text);
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
    }
}