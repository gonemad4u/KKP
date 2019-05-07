using DAL;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
namespace BLL
{
    public class OkuriJyo
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
               (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //Create the Excel file from dataset
        public void Okuri(DataTable ds,string SOKOCD)
        {
            string saveFileName = "送り状兼受領書_" + SOKOCD + "_" + DateTime.Now.ToString("yyyyMMddHHmmss");
            string templetFile = @"template\送り状兼受領書.xlsx";
            SaveFileDialog saveDialog = new SaveFileDialog {
                Filter = "Excel文件 |*.xlsx",
                FileName = saveFileName,
                RestoreDirectory = true
            };
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                saveFileName = saveDialog.FileName;
                File.Copy(templetFile, saveFileName, true);
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Open(saveFileName);
                Excel.Worksheet worksheet = workbook.Worksheets[1];
                excel.Visible = false;
                
                try
                {
                    log.Info("EXEC BEGIN");
                    var distinctIds = ds.AsEnumerable()
                     .Select(s => new
                     {
                         area = s.Field<string>("AREANM"),
                         soko = s.Field<string>("SOKONM"),
                         nukn = s.Field<string>("NUKNNM"),
                         
                     })
                       .Distinct().ToList();

                    if (distinctIds.Count > 1)
                    {
                        int no = 0;
                        foreach (var a in distinctIds)
                        {
                            string expression;
                            DataRow[] foundRows;
                            expression = "AREANM = '" + a.area.ToString() + "' AND SOKONM = '" + a.soko.ToString() + "'" + " AND NUKNNM = '" + a.nukn.ToString() + "'";

                            foundRows = ds.Select(expression);
                            worksheet.Name = foundRows[0][10] + "_" + foundRows[0][11] + "_" + foundRows[0][1];
                            workbook.Worksheets[1].Name = workbook.Worksheets[1].Name.Split(' ')[0];
                            worksheet.Cells[1, "H"] = foundRows[0][10];
                            worksheet.Cells[1, "AD"] = foundRows[0][11];
                            worksheet.Cells[5 , "B"] = foundRows[0][2];
                            worksheet.Cells[5 , "H"] = foundRows[0][5];
                            worksheet.Cells[5 , "N"] = foundRows[0][3];
                            worksheet.Cells[5 , "Y"] = foundRows[0][1];
                            worksheet.Cells[5 , "AS"] = foundRows[0][6];
                            worksheet.Cells[5 , "AW"] = foundRows[0][7];
                            worksheet.Cells[5 , "BC"] = foundRows[0][8];
                            worksheet.Cells[5 , "BM"] = foundRows[0][9];
                            worksheet.Cells[5 , "CD"] = foundRows[0][11];
                            worksheet.Cells[5, "CN"] = "サイン";

                            for (int i = 0; i < foundRows.Length-1; i++)
                            {
                                Excel.Range RngToCopy = worksheet.get_Range("B5").EntireRow;
                                Excel.Range RngToInsert = worksheet.get_Range("B" + (i + 6)).EntireRow;
                                RngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy());

                                worksheet.Cells[6 + i, "B"] = foundRows[i + 1][2];
                                worksheet.Cells[6 + i, "H"] = foundRows[i + 1][5];
                                worksheet.Cells[6 + i, "N"] = foundRows[i + 1][3];
                                worksheet.Cells[6 + i, "Y"] = foundRows[i + 1][1];
                                worksheet.Cells[6 + i, "AS"] = foundRows[i + 1][6];
                                worksheet.Cells[6 + i, "AW"] = foundRows[i + 1][7];
                                worksheet.Cells[6 + i, "BC"] = foundRows[i + 1][8];
                                worksheet.Cells[6 + i, "BM"] = foundRows[i + 1][9];
                                worksheet.Cells[6 + i, "CD"] = foundRows[i + 1][11];
                                worksheet.Cells[6 + i, "CN"] = "サイン";
                            }
                            no++;
                            if (no < distinctIds.Count)
                            {
                                worksheet.Copy(workbook.Worksheets[1], Type.Missing);
                                worksheet.get_Range("B5:B1000").EntireRow.ClearContents();
                                worksheet.get_Range("CN6:CN1000").EntireRow.Borders.LineStyle=Excel.XlLineStyle.xlLineStyleNone;
                            }
                        }
                    }

                    else
                    {
                        foreach (var a in distinctIds)
                        {
                            string expression;
                            DataRow[] foundRows;
                            expression = "AREANM = '" + a.area.ToString() + "' AND SOKONM = '" + a.soko.ToString() + "'" + " AND NUKNNM = '" + a.nukn.ToString() + "'";
                            foundRows = ds.Select(expression);
                            worksheet.Name = foundRows[0][10] + "_" + foundRows[0][11] + "_" + foundRows[0][1];
                            workbook.Worksheets[1].Name = workbook.Worksheets[1].Name.Split(' ')[0];
                            worksheet.Cells[1, "H"] = foundRows[0][10];
                            worksheet.Cells[1, "AD"] = foundRows[0][11];
                            worksheet.Cells[5 , "B"] = foundRows[0][2];
                            worksheet.Cells[5 , "H"] = foundRows[0][5];
                            worksheet.Cells[5 , "N"] = foundRows[0][3];
                            worksheet.Cells[5 , "Y"] = foundRows[0][1];
                            worksheet.Cells[5 , "AS"] = foundRows[0][6];
                            worksheet.Cells[5 , "AW"] = foundRows[0][7];
                            worksheet.Cells[5 , "BC"] = foundRows[0][8];
                            worksheet.Cells[5 , "BM"] = foundRows[0][9];
                            worksheet.Cells[5 , "CD"] = foundRows[0][11];
                            worksheet.Cells[5, "CN"] = "サイン";

                            for (int i = 1; i < foundRows.Length; i++)
                            {

                                Excel.Range RngToCopy = worksheet.get_Range("B5").EntireRow;
                                Excel.Range RngToInsert = worksheet.get_Range("B" + (i + 5)).EntireRow;
                                RngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy());

                                worksheet.Cells[5 + i, "B"] = foundRows[i][2];
                                worksheet.Cells[5 + i, "H"] = foundRows[i][5];
                                worksheet.Cells[5 + i, "N"] = foundRows[i][3];
                                worksheet.Cells[5 + i, "Y"] = foundRows[i][1];
                                worksheet.Cells[5 + i, "AS"] = foundRows[i][6];
                                worksheet.Cells[5 + i, "AW"] = foundRows[i][7];
                                worksheet.Cells[5 + i, "BC"] = foundRows[i][8];
                                worksheet.Cells[5 + i, "BM"] = foundRows[i][9];
                                worksheet.Cells[5 + i, "CD"] = foundRows[i][11];
                                worksheet.Cells[5 + i, "CN"] = "サイン";
                            }

                        }
                    }
                }
                catch (Exception e)
                {
                    log.Error("ERROR"+e.Message );
                }

                finally
                {
                    excel.DisplayAlerts = false;
                    Application.DoEvents();
                    workbook.Save();
                    excel.Quit();
                    Marshal.FinalReleaseComObject(excel);

                    log.Info("EXEC END");
                }
            }
            else
            {
                MessageBox.Show("処理を中止しました。");
            }
        }

        //Search database from the information from interface
        public void SearchParameter(DataGridView gr)
        {
            try

            {
                DEV10G2U dev = new DEV10G2U();
                StringBuilder sql = new StringBuilder();

                sql.Append("SELECT                                   ");                
                sql.Append("T_KDHSINFO.AREACD,                       ");
                sql.Append("T_KDHSINFO.NUKNNM,                       ");
                sql.Append("T_KDHSINFO.SYUKABI,                      ");
                sql.Append("T_KDHSINFO.DENPYONO,                     ");
                sql.Append("T_KDHSINFO.IRIGSYCD,                     ");
                sql.Append("T_KDHSINFO.NOUKIBI,                      ");               
                sql.Append("T_KDHSINFO.CHIKUCD,                      ");
                sql.Append("T_KDHSINFO.KOSU,                         ");
                sql.Append("T_KDHSINFO.WT,                           ");
                sql.Append("T_KDHSINFO.BIKO,                         ");
                sql.Append("AREA.AREANM,                           ");
                sql.Append("SOKO.SOKONM                            ");
                sql.Append("FROM                                     ");
                sql.Append("T_KDHSINFO                               ");
                sql.Append("INNER JOIN    (SELECT DISTINCT AREACD, AREANM FROM M_AREA)AREA                       ");
                sql.Append("ON   AREA.AREACD = T_KDHSINFO.AREACD   ");
                sql.Append("INNER JOIN  (SELECT DISTINCT SOKOCD, SOKONM FROM M_SOKO)SOKO                       ");
                sql.Append("ON   SOKO.SOKOCD = T_KDHSINFO.IRIGSYCD ");
                sql.Append("    WHERE                                ");
                string denpyonoSql = "T_KDHSINFO.DENPYONO IN(      ";
                string sykFileNMSql = "AND T_KDHSINFO.SYKFILENM IN (     ";
                string seqNoSql = "AND T_KDHSINFO.SEQNO IN (     ";

                for (int i = 0; i < gr.RowCount; i++)
                {
                    if ((bool)gr.Rows[i].Cells[0].EditedFormattedValue == true)
                    {
                        denpyonoSql += "'" + gr.Rows[i].Cells[5].Value + "'" + ",";
                        sykFileNMSql += "'" + gr.Rows[i].Cells["SYKFILENM"].Value + "'" + ",";
                        seqNoSql += gr.Rows[i].Cells["SEQNO"].Value + ",";

                    }
                }
                denpyonoSql = denpyonoSql.Remove(denpyonoSql.Length - 1);
                sql.Append(denpyonoSql + ") ");
                sykFileNMSql = sykFileNMSql.Remove(sykFileNMSql.Length - 1);
                sql.Append(sykFileNMSql + ") ");
                seqNoSql = seqNoSql.Remove(seqNoSql.Length - 1);
                sql.Append(seqNoSql + ") ");

                
                var SOKOCD = Tools.getSokocd();

                sql.Append("    AND T_KDHSINFO.SOKOCD =" + "'"+SOKOCD + "'");

                DataTable dataSet = dev.executeSelectQuery(sql.ToString());

                Okuri(dataSet, SOKOCD);

            }
            catch

            {
                throw;

            }
        }

      
    }

    
}
