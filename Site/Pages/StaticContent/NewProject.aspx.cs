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
    public partial class NewProject : System.Web.UI.Page
    {
        protected int ProjectId;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (IsPostBack)
            {
                
            }
        }

        protected void saveProject_Click(object sender, EventArgs e)
        {

            string projectName = "Project";
            var date = DateTime.Now;
            string status = "Begin";

            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = "insert into siterisk.projects(projectName,modified,time,status) values(@projectName,@modified,@time, @status); SELECT LAST_INSERT_ID();";
                string count = "Select count(*) from siterisk.projects";

                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };
                MySqlCommand cmd2 = new MySqlCommand(count)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };
                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@projectName", projectName + Convert.ToInt32(cmd2.ExecuteScalar()).ToString());
                cmd.Parameters.AddWithValue("@modified", date);
                cmd.Parameters.AddWithValue("@time", 0);
                cmd.Parameters.AddWithValue("@status", status);
                ProjectId = Convert.ToInt32(cmd.ExecuteScalar());
                Session["projectId"] = ProjectId;
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
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = "insert into siterisk.accountinfo(FullName,ShortName,Address, ActualAddress, Contact, Phone, Email, DamagePrice, projectId) " +
                    "values(@FullName,@ShortName, @Address, @ActualAddress, @Contact, @Phone, @Email, @DamagePrice, @projectId)";

                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };
                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@FullName", accountName.Text);
                cmd.Parameters.AddWithValue("@ShortName", shortNameAccountName.Text);
                cmd.Parameters.AddWithValue("@Address", accountAddress.Text);
                cmd.Parameters.AddWithValue("@ActualAddress", accountActualAdress.Text);
                cmd.Parameters.AddWithValue("@Contact", contact.Text);
                cmd.Parameters.AddWithValue("@Phone", phone.Text);
                cmd.Parameters.AddWithValue("@Email", email.Text);
                cmd.Parameters.AddWithValue("@DamagePrice", Convert.ToDouble(normDamage.Text));
                cmd.Parameters.AddWithValue("@projectId", Convert.ToInt32(Session["projectId"]));
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
            Response.Redirect($"StepTables/Stage2.aspx?projectId={Convert.ToInt32(Session["projectId"]).ToString()}");
        }
    }
}