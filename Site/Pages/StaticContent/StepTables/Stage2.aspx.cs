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
    public partial class Stage2 : System.Web.UI.Page
    {
        private static DataTable dataTable;

        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                dataTable = new DataTable();
                dataTable.Columns.Add("threatId");
                dataTable.Columns.Add("threatName");
                dataTable.Columns.Add("threatFullName");
                PageDataBind();
            }
        }

        protected void threatGrid_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {

        }

        protected void PageDataBind()
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
                RadioButtonList.DataSource = dv;
                RadioButtonList.DataTextField = "threatName";
                RadioButtonList.DataValueField = "threatId";
                RadioButtonList.DataBind();

                threatGrid.DataSource = dataTable;
                threatGrid.DataBind();
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

        protected void RadioButtonList_SelectedIndexChanged(object sender, EventArgs e)
        {
            
        }

        protected void saveProject_Click(object sender, EventArgs e)
        {
            
        }

        protected void continueProject_Click(object sender, EventArgs e)
        {
            Response.Redirect($"Stage3.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void threatGrid_RowCommand(object sender, GridViewCommandEventArgs e)
        {
            int index = Convert.ToInt32(e.CommandArgument);
            if (e.CommandName == "AddLink")
            {
                MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
                try
                {
                    string insertQuery = "insert into siterisk.actualthreats(Name,FullName,Description,FSTEKId,projectId) values(@Name,@FullName,@Description,@FSTEKId,@projectId)";


                    MySqlCommand cmd = new MySqlCommand(insertQuery)
                    {
                        Connection = con,
                        CommandType = CommandType.Text
                    };

                    con.Open();
                    cmd.Prepare();


                    cmd.Parameters.AddWithValue("@Name", threatGrid.Rows[index].Cells[1].Text);
                    cmd.Parameters.AddWithValue("@FullName", threatGrid.Rows[index].Cells[2].Text);
                    cmd.Parameters.AddWithValue("@Description", "");
                    cmd.Parameters.AddWithValue("@FSTEKId", threatGrid.Rows[index].Cells[0].Text);
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
        }
    }
}