﻿using MySql.Data.MySqlClient;
using System;
using System.Data;

namespace Site.Pages.StaticContent.StepTables
{
    public partial class Stage5 : System.Web.UI.Page
    {
        private static DataTable dataTable;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                dataTable = new DataTable();
                dataTable.Columns.Add("assetId");
                dataTable.Columns.Add("assetName");
                PageDataBind();
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
                string insertQuery = "insert into siterisk.threateffectib(Name,Effect, assetType,projectId) values(@Name,@Effect, @assetType,@projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@Name", IAName.Text);
                cmd.Parameters.AddWithValue("@Effect", IADescription.Text);
                cmd.Parameters.AddWithValue("@assetType", IAType.SelectedItem.Text);
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

        protected void continueProject_Click(object sender, EventArgs e)
        {
            Response.Redirect($"Stage6.aspx?projectId={Request.QueryString["projectId"]}");
        }
    }
}