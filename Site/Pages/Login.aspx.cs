using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Site.Pages
{
    public partial class Login : System.Web.UI.Page
    {
        private DataTable dataTable;
        protected void Page_Load(object sender, EventArgs e)
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("id");
            dataTable.Columns.Add("login");
            dataTable.Columns.Add("password");
        }

        protected void login_Click(object sender, EventArgs e)
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * From siterisk.users where login=@login && password=@password";
                MySqlCommand cmd = new MySqlCommand(selectquery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };
                con.Open();
                cmd.Parameters.AddWithValue("@login", loginTexBox.Text);
                cmd.Parameters.AddWithValue("@password", passwordTextBox.Text);
                cmd.ExecuteNonQuery();

                MySqlDataReader rdr = cmd.ExecuteReader();

                while (rdr.Read())
                {
                    DataRow dr = dataTable.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["login"] = rdr.GetValue(rdr.GetOrdinal("login"));
                    dr["password"] = rdr.GetValue(rdr.GetOrdinal("password"));

                    dataTable.Rows.Add(dr);
                }

            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                con.Close();
            }
            if (dataTable.Rows.Count == 1)
            {
                Response.Redirect("Projects.aspx");
            }
            else 
            {
                ClientScript.RegisterStartupScript(this.GetType(), "myalert", "alert('Invalid user');", true);
            }
            
        }
    }
}