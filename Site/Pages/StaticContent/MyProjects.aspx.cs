using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Site.Pages.StaticContent
{
    public partial class MyProjects : System.Web.UI.Page
    {
        private static DataTable dataTable;

        protected void Page_Load(object sender, EventArgs e)
        {

            if (!IsPostBack)
            {
                dataTable = new DataTable();
                dataTable.Columns.Add("ProjectId");
                dataTable.Columns.Add("DataChange");
                dataTable.Columns.Add("TimeWorking");
                dataTable.Columns.Add("Status");
                PageDataBind();

            }
        }

        protected void PageDataBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = "SELECT * FROM siterisk.projects";
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
                    dr["ProjectId"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["DataChange"] = rdr.GetValue(rdr.GetOrdinal("projectName"));
                    dr["TimeWorking"] = rdr.GetValue(rdr.GetOrdinal("time"));
                    dr["Status"] = rdr.GetValue(rdr.GetOrdinal("status"));


                    dataTable.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable);
                MyProjectsGrid.DataSource = dv;
                MyProjectsGrid.DataBind();

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

        protected void MyProjectsGrid_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            int index = Convert.ToInt32(e.CommandArgument);
            if (e.CommandName == "UpdateProject")
            {
                ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('Update is successfull')", true);
            }

            if (e.CommandName == "Download")
            {
                String cellText = MyProjectsGrid.Rows[index].Cells[3].Text;
                if (cellText != "Done")
                {
                    MyProjectsGrid.Rows[index].Cells[4].Enabled = false;
                }

                ClientScript.RegisterStartupScript(this.GetType(), "alert", "alert('Begin downloading...')", true);
            }
        }
    }
}