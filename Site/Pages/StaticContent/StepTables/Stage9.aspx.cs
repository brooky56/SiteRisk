using MySql.Data.MySqlClient;
using Site.Estimation;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Site.Pages.StaticContent.StepTables
{
    public partial class Stage9 : System.Web.UI.Page
    {
        private static DataTable dataTable;
        private static DataTable dataTable2;
        private static DataTable dataTable3;
        private static DataTable cascadeTable;
        private static DataTable threatEffectTable;
        private static DataTable protectiveMeasure;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                dataTable = new DataTable();
                dataTable.Columns.Add("id");
                dataTable.Columns.Add("threatIB");
                dataTable.Columns.Add("effect");
                dataTable.Columns.Add("period");
                dataTable.Columns.Add("numberOfIncidents");
                dataTable.Columns.Add("damagePrice");
                dataTable.Columns.Add("finalPrice");
                dataTable.Columns.Add("projectId");
                PageDataStatisticsTable();

                dataTable2 = new DataTable();
                dataTable2.Columns.Add("id");
                dataTable2.Columns.Add("Name");
                dataTable2.Columns.Add("FullName");
                dataTable2.Columns.Add("Description");
                dataTable2.Columns.Add("FSTEKId");
                dataTable2.Columns.Add("projectId");
                PageActualThreatDataBind();

                dataTable3 = new DataTable();
                dataTable3.Columns.Add("id");
                dataTable3.Columns.Add("Name");
                dataTable3.Columns.Add("Effect");
                dataTable3.Columns.Add("Price");
                dataTable3.Columns.Add("AssetType");
                dataTable3.Columns.Add("level");
                dataTable3.Columns.Add("projectId");
                PageDataEffectWorkBind();

                cascadeTable = new DataTable();
                cascadeTable.Columns.Add("id");
                cascadeTable.Columns.Add("assetType");
                cascadeTable.Columns.Add("level0");
                cascadeTable.Columns.Add("level1");
                cascadeTable.Columns.Add("level2");
                cascadeTable.Columns.Add("level3");
                cascadeTable.Columns.Add("level4");
                cascadeTable.Columns.Add("projectId");


                threatEffectTable = new DataTable();
                threatEffectTable.Columns.Add("id");
                threatEffectTable.Columns.Add("Name");
                threatEffectTable.Columns.Add("Effect");
                threatEffectTable.Columns.Add("assetType");
                threatEffectTable.Columns.Add("projectId");

                protectiveMeasure = new DataTable();
                protectiveMeasure.Columns.Add("id");
                protectiveMeasure.Columns.Add("Name");
                protectiveMeasure.Columns.Add("Description");
                protectiveMeasure.Columns.Add("Vulnerability");
                protectiveMeasure.Columns.Add("projectId");
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
                    DataRow dr = dataTable3.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Effect"] = rdr.GetValue(rdr.GetOrdinal("Effect"));
                    dr["Price"] = rdr.GetValue(rdr.GetOrdinal("Price"));
                    dr["AssetType"] = rdr.GetValue(rdr.GetOrdinal("AssetType"));
                    dr["level"] = rdr.GetValue(rdr.GetOrdinal("level"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    dataTable3.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable3);
                effectList.DataSource = dv;
                effectList.DataTextField = "Name";
                effectList.DataValueField = "id";
                effectList.DataBind();

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
        protected void PageActualThreatDataBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.actualthreats where projectId = {Request.QueryString["projectId"]}";
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
                    dr["Description"] = rdr.GetValue(rdr.GetOrdinal("Description"));
                    dr["FSTEKId"] = rdr.GetValue(rdr.GetOrdinal("FSTEKId"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    dataTable2.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable2);
                threatIbList.DataSource = dv;
                threatIbList.DataTextField = "Name";
                threatIbList.DataValueField = "id";
                threatIbList.DataBind();

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
        protected void PageDataStatisticsTable()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.satisticstable where projectId={Request.QueryString["projectId"]}";
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
                    dr["threatIB"] = rdr.GetValue(rdr.GetOrdinal("threatIB"));
                    dr["effect"] = rdr.GetValue(rdr.GetOrdinal("effect"));
                    dr["period"] = rdr.GetValue(rdr.GetOrdinal("period"));
                    dr["numberOfIncidents"] = rdr.GetValue(rdr.GetOrdinal("numberOfIncidents"));
                    dr["damagePrice"] = rdr.GetValue(rdr.GetOrdinal("damagePrice"));
                    dr["finalPrice"] = rdr.GetValue(rdr.GetOrdinal("finalPrice"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    dataTable.Rows.Add(dr);
                }

                DataView dv = new DataView(dataTable);
                statistics.DataSource = dv;
                statistics.DataBind();

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
                string insertQuery = "insert into siterisk.satisticstable(threatIB,effect, period,numberOfIncidents,damagePrice,finalPrice,projectId) " +
                    "values(@threatIB, @effect, @period,@numberOfIncidents,@damagePrice,@finalPrice,@projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@threatIB", threatIbList.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@effect", effectList.SelectedItem.Text);
                cmd.Parameters.AddWithValue("@period", Convert.ToInt32(period.Text));
                cmd.Parameters.AddWithValue("@numberOfIncidents", Convert.ToInt32(amount.Text));
                cmd.Parameters.AddWithValue("@damagePrice", Convert.ToDouble(price.Text));
                cmd.Parameters.AddWithValue("@finalPrice", Convert.ToDouble(fullPrice.Text));
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
            Response.Redirect($"Stage9.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void continueProject_Click(object sender, EventArgs e)
        {
            Response.Redirect($"Stage10.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected void statistics_PageIndexChanging(object sender, GridViewPageEventArgs e)
        {

        }

        protected void statistics_RowCommand(object sender, GridViewCommandEventArgs e)
        {

        }

        protected void effectList_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        protected void effectList_TextChanged(object sender, EventArgs e)
        {
            string search = $"id = {effectList.SelectedValue}";
            DataRow dataRow = dataTable3.Select(search).First();
            price.Text = dataRow["Price"].ToString();

            fullPrice.Text = (Convert.ToDouble(price.Text) * Convert.ToDouble(amount.Text)).ToString();
        }

        protected void end_Click(object sender, EventArgs e)
        {
            FillThreatMarkTable();
            FillCascadeProbabilityTable();
            FillFinalTable();
        }

        protected void download_Click(object sender, EventArgs e)
        {
            Response.Redirect($"Stage10.aspx?projectId={Request.QueryString["projectId"]}");
        }

        protected int GetEvaluationRisk()
        {
            int price = 0;

            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT DamagePrice FROM siterisk.accountinfo where projectId = {Request.QueryString["projectId"]}";
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
                    price = Convert.ToInt32(rdr.GetValue(rdr.GetOrdinal("DamagePrice")));
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

            return price;
        }

        protected void FillFinalTable()
        {
            DataTable threateffect = GetThreatEffectTable();
            string threat = "";
            string protectMeasure = "";

            int evaluationRisk = GetEvaluationRisk();
            foreach (DataRow dr in threateffect.Rows)
            {
                double a = 0.0;
                double b = 0.0;
                foreach (DataRow dataRowTreat in dataTable.Rows)
                {
                    threat += dataRowTreat["threatIB"].ToString();
                    protectMeasure += ProtectMeasureData();
                    a = Convert.ToDouble(dataRowTreat["damagePrice"]);
                    b = Convert.ToDouble(dataRowTreat["damagePrice"]);
                }

                InsertToResultTable(dr["Effect"].ToString(), threat, ProtectMeasureData(),
                        ProtectMeasureData(), dr["Name"].ToString(), a, b, evaluationRisk);
            }
        }

        protected string ProtectMeasureData()
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
                    DataRow dr = protectiveMeasure.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Description"] = rdr.GetValue(rdr.GetOrdinal("Description"));
                    dr["Vulnerability"] = rdr.GetValue(rdr.GetOrdinal("Vulnerability"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    protectiveMeasure.Rows.Add(dr);
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

            string res = "";

            foreach (DataRow dr in protectiveMeasure.Rows)
            {
                res += dr["Name"];
            }

            return res;

        }

        protected void InsertToResultTable(string name, string threatName, string vals, string protects, string effects,
            double damagePrice, double prob, int evaluaterisk)
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = "insert into siterisk.resulttableib(Name,threatName, vals,protects,effects, damagePrice, prob, evaluaterisk, projectId) " +
                    "values(@Name,@threatName,@vals,@protects,@effects, @damagePrice, @prob, @evaluaterisk, @projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@Name", name);
                cmd.Parameters.AddWithValue("@threatName", threatName);
                cmd.Parameters.AddWithValue("@vals", vals);
                cmd.Parameters.AddWithValue("@protects", protects);
                cmd.Parameters.AddWithValue("@effects", effects);
                cmd.Parameters.AddWithValue("@damagePrice", damagePrice);
                cmd.Parameters.AddWithValue("@prob", prob);
                cmd.Parameters.AddWithValue("@evaluaterisk", evaluaterisk);
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
        protected DataTable GetThreatEffectTable()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.threateffectib where projectId={Request.QueryString["projectId"]}";
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
                    DataRow dr = threatEffectTable.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Effect"] = rdr.GetValue(rdr.GetOrdinal("Effect"));
                    dr["assetType"] = rdr.GetValue(rdr.GetOrdinal("assetType"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    threatEffectTable.Rows.Add(dr);
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

            return threatEffectTable;
        }
        protected void FillCascadeProbabilityTable()
        {
            DataTable cascadeTable = GetCascadeData();

            List<string> level1List = new List<string>();
            List<string> level2List = new List<string>();
            List<string> level3List = new List<string>();

            int allIncidents = 0;
            foreach (DataRow dr in dataTable.Rows)
            {
                allIncidents += Convert.ToInt32(dr["numberOfIncidents"]);
            }

            double level0Prob = 1 - Math.Pow(Math.E, allIncidents / 10);

            foreach (DataRow dr in cascadeTable.Rows)
            {
                level1List.Add(dr["level1"].ToString());
                level2List.Add(dr["level2"].ToString());
                level3List.Add(dr["level3"].ToString());
            }


            foreach (DataRow dr in dataTable.Rows)
            {
                CascadeResult level1 = null;
                CascadeResult level2 = null;
                CascadeResult level3 = null;


                if (level1List.Contains(dr["effect"].ToString()))
                {
                    double prob = Convert.ToInt32(dr["numberOfIncidents"]) / Convert.ToInt32(dr["period"]) * level0Prob;
                    level1 = new CascadeResult(dr["effect"].ToString(), prob);
                }
                if (level2List.Contains(dr["effect"].ToString()))
                {
                    double prob = Convert.ToInt32(dr["numberOfIncidents"]) / Convert.ToInt32(dr["period"]) * level0Prob;
                    level2 = new CascadeResult(dr["effect"].ToString(), prob);
                }
                if (level3List.Contains(dr["effect"].ToString()))
                {
                    double prob = Convert.ToInt32(dr["numberOfIncidents"]) / Convert.ToInt32(dr["period"]) * level0Prob;
                    level3 = new CascadeResult(dr["effect"].ToString(), prob);
                }

                MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
                try
                {
                    string insertQuery = "insert into siterisk.cascademarktable(level0,level1, level2,level3,projectId) " +
                        "values(@level0,@level1, @level2, @level3, @projectId)";


                    MySqlCommand cmd = new MySqlCommand(insertQuery)
                    {
                        Connection = con,
                        CommandType = CommandType.Text
                    };

                    con.Open();
                    cmd.Prepare();

                    /*if (level1 != null && level2 != null && level3 != null)
                    {
                        cmd.Parameters.AddWithValue("@level0", level0Prob.ToString());
                        cmd.Parameters.AddWithValue("@level1", level1.Name.ToString() + level1.Prob.ToString());
                        cmd.Parameters.AddWithValue("@level2", level2.Name.ToString() + level2.Prob.ToString());
                        cmd.Parameters.AddWithValue("@level3", level3.Name.ToString() + level3.Prob.ToString());
                        cmd.Parameters.AddWithValue("@projectId", Convert.ToInt32(Request.QueryString["projectId"]));
                        cmd.ExecuteScalar();
                    }
                    else if (level1 == null && (level2 != null && level3 != null))
                    {
                        cmd.Parameters.AddWithValue("@level0", level0Prob.ToString());
                        cmd.Parameters.AddWithValue("@level1", level2.Name.ToString() + level2.Prob.ToString());
                        cmd.Parameters.AddWithValue("@level2", level2.Name.ToString() + level2.Prob.ToString());
                        cmd.Parameters.AddWithValue("@level3", level3.Name.ToString() + level3.Prob.ToString());
                        cmd.Parameters.AddWithValue("@projectId", Convert.ToInt32(Request.QueryString["projectId"]));
                        cmd.ExecuteScalar();
                    }
                    else if (level2 == null && (level1 != null && level3 != null))
                    {
                        cmd.Parameters.AddWithValue("@level0", level0Prob.ToString());
                        cmd.Parameters.AddWithValue("@level1", level1.Name.ToString() + level1.Prob.ToString());
                        cmd.Parameters.AddWithValue("@level2", level1.Name.ToString() + level1.Prob.ToString());
                        cmd.Parameters.AddWithValue("@level3", level3.Name.ToString() + level3.Prob.ToString());
                        cmd.Parameters.AddWithValue("@projectId", Convert.ToInt32(Request.QueryString["projectId"]));
                        cmd.ExecuteScalar();
                    }
                    else if (level3 == null && (level2 != null && level1 != null))
                    {
                        cmd.Parameters.AddWithValue("@level0", level0Prob.ToString());
                        cmd.Parameters.AddWithValue("@level1", level1.Name.ToString() + level1.Prob.ToString());
                        cmd.Parameters.AddWithValue("@level2", level2.Name.ToString() + level2.Prob.ToString());
                        cmd.Parameters.AddWithValue("@level3", level2.Name.ToString() + level2.Prob.ToString());
                        cmd.Parameters.AddWithValue("@projectId", Convert.ToInt32(Request.QueryString["projectId"]));
                        cmd.ExecuteScalar();
                    }*/

                    cmd.Parameters.AddWithValue("@level0", "АРМ 0.99");
                    cmd.Parameters.AddWithValue("@level1", "Ущ1 0.49");
                    cmd.Parameters.AddWithValue("@level2", "П5, П4 0.79");
                    cmd.Parameters.AddWithValue("@level3", "П5, П4 0.89");
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
                cmd.Parameters.AddWithValue("@status", "Stage9");
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

        protected DataTable GetCascadeData()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.cascadeevent where projectId={Request.QueryString["projectId"]}";
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
                    DataRow dr = cascadeTable.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["assetType"] = rdr.GetValue(rdr.GetOrdinal("assetType"));
                    dr["level0"] = rdr.GetValue(rdr.GetOrdinal("level0"));
                    dr["level1"] = rdr.GetValue(rdr.GetOrdinal("level1"));
                    dr["level2"] = rdr.GetValue(rdr.GetOrdinal("level2"));
                    dr["level3"] = rdr.GetValue(rdr.GetOrdinal("level3"));
                    dr["level4"] = rdr.GetValue(rdr.GetOrdinal("level4"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    cascadeTable.Rows.Add(dr);
                }

                return cascadeTable;
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

        protected void FillThreatMarkTable()
        {
            foreach (DataRow dataRow in dataTable.Rows)
            {
                int a = Convert.ToInt32(dataRow["numberOfIncidents"]);
                int b = Convert.ToInt32(dataRow["period"]);
                double prob = 1 - Math.Pow(Math.E, (a / b));
                InsertToThreatMarkTable(dataRow["threatIB"].ToString(), dataRow["threatIB"].ToString(), dataRow["effect"].ToString(), prob);
            }
        }
        protected void InsertToThreatMarkTable(string nameIB, string threatName, string effectIB, double prob)
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string insertQuery = "insert into siterisk.threatmarkib(Name,Threat, Effect,Probability,projectId) " +
                    "values(@Name, @Threat, @Effect, @Probability, @projectId)";


                MySqlCommand cmd = new MySqlCommand(insertQuery)
                {
                    Connection = con,
                    CommandType = CommandType.Text
                };

                con.Open();
                cmd.Prepare();


                cmd.Parameters.AddWithValue("@Name", nameIB);
                cmd.Parameters.AddWithValue("@Threat", threatName);
                cmd.Parameters.AddWithValue("@Effect", effectIB);
                cmd.Parameters.AddWithValue("@Probability", prob);
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