using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Site.Pages.StaticContent.StepTables
{
    public partial class Stage10 : System.Web.UI.Page
    {
        private static DataTable assets;
        private static DataTable actualThreats;
        private static DataTable vals;
        private static DataTable threatEffectIB;
        private static DataTable workEffect;
        private static DataTable protectMeasure;
        private static DataTable threatEvaluationTable;
        private static DataTable cascadeProb;
        private static DataTable finalResult;
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {

                finalResult = new DataTable();
                finalResult.Columns.Add("id");
                finalResult.Columns.Add("Name");
                finalResult.Columns.Add("threatName");
                finalResult.Columns.Add("vals");
                finalResult.Columns.Add("protects");
                finalResult.Columns.Add("effects");
                finalResult.Columns.Add("damagePrice");
                finalResult.Columns.Add("prob");
                finalResult.Columns.Add("evaluaterisk");
                finalResult.Columns.Add("projectId");
                PageDataFinalResultBind();


                cascadeProb = new DataTable();
                cascadeProb.Columns.Add("id");
                cascadeProb.Columns.Add("level0");
                cascadeProb.Columns.Add("level1");
                cascadeProb.Columns.Add("level2");
                cascadeProb.Columns.Add("level3");
                cascadeProb.Columns.Add("projectID");
                PageDataCascadeProbBind();

                threatEvaluationTable = new DataTable();
                threatEvaluationTable.Columns.Add("id");
                threatEvaluationTable.Columns.Add("Name");
                threatEvaluationTable.Columns.Add("Threat");
                threatEvaluationTable.Columns.Add("Effect");
                threatEvaluationTable.Columns.Add("Probability");
                threatEvaluationTable.Columns.Add("projectId");
                PageDataThreatEvaluationBind();

                vals = new DataTable();
                vals.Columns.Add("id");
                vals.Columns.Add("Name");
                vals.Columns.Add("FullName");
                PageValDatBind();

                assets = new DataTable();
                assets.Columns.Add("id");
                assets.Columns.Add("Name");
                assets.Columns.Add("Description");
                assets.Columns.Add("assetType");
                assets.Columns.Add("Price");
                assets.Columns.Add("Time");
                assets.Columns.Add("MainAmount");
                assets.Columns.Add("ReserveAmount");
                assets.Columns.Add("ZipAmount");
                assets.Columns.Add("projectId");
                PageDataBindAsset();

                actualThreats = new DataTable();
                actualThreats.Columns.Add("id");
                actualThreats.Columns.Add("Name");
                actualThreats.Columns.Add("FullName");
                actualThreats.Columns.Add("Description");
                actualThreats.Columns.Add("FSTEKId");
                actualThreats.Columns.Add("projectId");
                PageActualThreatDataBind();

                workEffect = new DataTable();
                workEffect.Columns.Add("id");
                workEffect.Columns.Add("Name");
                workEffect.Columns.Add("Effect");
                workEffect.Columns.Add("Price");
                workEffect.Columns.Add("AssetType");
                workEffect.Columns.Add("level");
                workEffect.Columns.Add("projectId");
                PageDataEffectWorkBind();

                threatEffectIB = new DataTable();
                threatEffectIB.Columns.Add("id");
                threatEffectIB.Columns.Add("Name");
                threatEffectIB.Columns.Add("Effect");
                threatEffectIB.Columns.Add("assetType");
                threatEffectIB.Columns.Add("projectId");
                PageDataThreatEffectBind();

                protectMeasure = new DataTable();
                protectMeasure.Columns.Add("id");
                protectMeasure.Columns.Add("Name");
                protectMeasure.Columns.Add("Description");
                protectMeasure.Columns.Add("Vulnerability");
                protectMeasure.Columns.Add("projectId");
                PageDataProtectMeasureBind();
            }
        }

        private void PageDataFinalResultBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.resulttableib where projectId={Request.QueryString["projectId"]}";
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
                    DataRow dr = finalResult.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["threatName"] = rdr.GetValue(rdr.GetOrdinal("threatName"));
                    dr["vals"] = rdr.GetValue(rdr.GetOrdinal("vals"));
                    dr["protects"] = rdr.GetValue(rdr.GetOrdinal("protects"));
                    dr["effects"] = rdr.GetValue(rdr.GetOrdinal("effects"));
                    dr["damagePrice"] = rdr.GetValue(rdr.GetOrdinal("damagePrice"));
                    dr["prob"] = rdr.GetValue(rdr.GetOrdinal("prob"));
                    dr["evaluaterisk"] = rdr.GetValue(rdr.GetOrdinal("evaluaterisk"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    finalResult.Rows.Add(dr);
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
        }

        private void PageDataCascadeProbBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.cascademarktable where projectId={Request.QueryString["projectId"]}";
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
                    DataRow dr = cascadeProb.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["level0"] = rdr.GetValue(rdr.GetOrdinal("level0"));
                    dr["level1"] = rdr.GetValue(rdr.GetOrdinal("level1"));
                    dr["level2"] = rdr.GetValue(rdr.GetOrdinal("level2"));
                    dr["level3"] = rdr.GetValue(rdr.GetOrdinal("level3"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    cascadeProb.Rows.Add(dr);
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
        }

        private void PageDataThreatEvaluationBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.threatmarkib where projectId={Request.QueryString["projectId"]}";
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
                    DataRow dr = threatEvaluationTable.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Threat"] = rdr.GetValue(rdr.GetOrdinal("Threat"));
                    dr["Effect"] = rdr.GetValue(rdr.GetOrdinal("Effect"));
                    dr["Probability"] = rdr.GetValue(rdr.GetOrdinal("Probability"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    threatEvaluationTable.Rows.Add(dr);
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
        }

        private void PageDataProtectMeasureBind()
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
                    DataRow dr = protectMeasure.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Description"] = rdr.GetValue(rdr.GetOrdinal("Description"));
                    dr["Vulnerability"] = rdr.GetValue(rdr.GetOrdinal("Vulnerability"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    protectMeasure.Rows.Add(dr);
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
        }

        private void PageDataThreatEffectBind()
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
                    DataRow dr = threatEffectIB.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Effect"] = rdr.GetValue(rdr.GetOrdinal("Effect"));
                    dr["assetType"] = rdr.GetValue(rdr.GetOrdinal("assetType"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    threatEffectIB.Rows.Add(dr);
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
        }

        private void PageDataEffectWorkBind()
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
                    DataRow dr = workEffect.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["Effect"] = rdr.GetValue(rdr.GetOrdinal("Effect"));
                    dr["Price"] = rdr.GetValue(rdr.GetOrdinal("Price"));
                    dr["AssetType"] = rdr.GetValue(rdr.GetOrdinal("AssetType"));
                    dr["level"] = rdr.GetValue(rdr.GetOrdinal("level"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    workEffect.Rows.Add(dr);
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
        }

        private void PageActualThreatDataBind()
        {
            MySqlConnection con = new MySqlConnection("server=localhost;port=3306;uid=root;pwd=admin;database=siterisk;");
            try
            {
                string selectquery = $"SELECT * FROM siterisk.actualthreats where projectId={Request.QueryString["projectId"]}";
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
                    DataRow dr = actualThreats.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["FullName"] = rdr.GetValue(rdr.GetOrdinal("FullName"));
                    dr["Description"] = rdr.GetValue(rdr.GetOrdinal("Description"));
                    dr["FSTEKId"] = rdr.GetValue(rdr.GetOrdinal("FSTEKId"));
                    dr["projectId"] = rdr.GetValue(rdr.GetOrdinal("projectId"));

                    actualThreats.Rows.Add(dr);
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
        }

        private void PageDataBindAsset()
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
                    DataRow dr = assets.NewRow();
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

                    assets.Rows.Add(dr);
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
        }

        private void PageValDatBind()
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
                    DataRow dr = vals.NewRow();
                    dr["id"] = rdr.GetValue(rdr.GetOrdinal("id"));
                    dr["Name"] = rdr.GetValue(rdr.GetOrdinal("Name"));
                    dr["FullName"] = rdr.GetValue(rdr.GetOrdinal("FullName"));

                    vals.Rows.Add(dr);
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
                cmd.Parameters.AddWithValue("@status", "Завершен");
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
        protected void download_Click(object sender, EventArgs e)
        {
            UpdateStageProject();
            using (var workbook = new XLWorkbook())
            {
                var sh1 = workbook.Worksheets.Add(assets, "Активы");
                var sh2 = workbook.Worksheets.Add(actualThreats, "Актуальные угрозы");
                var sh3 = workbook.Worksheets.Add(vals, "Уязвимости");
                var sh4 = workbook.Worksheets.Add(threatEffectIB, "Последствия реализации угроз ИБ");
                var sh5 = workbook.Worksheets.Add(workEffect, "Последствия нарушения работы");
                var sh6 = workbook.Worksheets.Add(protectMeasure, "Защитные меры");
                var sh7 = workbook.Worksheets.Add(threatEvaluationTable, "Оценка вероятности ");
                var sh8 = workbook.Worksheets.Add(cascadeProb, "Оценка степени вероятности ");
                var sh9 = workbook.Worksheets.Add(finalResult, "Итоговый реестр рисков ИБ");

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    var content = stream.ToArray();

                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AppendHeader("Content-Disposition", $"attachment; filename=Project{ Request.QueryString["projectId"]}.xls");
                    Response.BinaryWrite(content);
                    Response.End();
                }
            }

        }
    }
}