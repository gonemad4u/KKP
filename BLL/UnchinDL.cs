using DAL;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
namespace BLL
{
    public class UnchinDl
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
        (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void UnchinDownLoad(DataSet ds, DataGridView gr)
        {

            string saveFileName = "運賃計算_" + DateTime.Now.ToString("yyyyMMdd");
            string templetFile = @"template\運賃計算.xlsx";
            SaveFileDialog saveDialog = new SaveFileDialog
            {
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
                List<string[]> vs = new List<string[]>();//Create a sring list that stores the four keys
                int excelRow = 0;
                int number = 0;
                int currentRow = 2;
                try
                {
                    log.Info("EXEC BEGIN");
                    for (int i = 0; i < gr.RowCount; i++)
                    {
                        if ((bool)gr.Rows[i].Cells[0].EditedFormattedValue == true)
                        {
                            vs.Add(new string[]{ gr.Rows[i].Cells["SOKOCD"].Value.ToString(),
                        gr.Rows[i].Cells["SYKFILENM"].Value.ToString(),
                        gr.Rows[i].Cells["SEQNO"].Value.ToString(),
                        gr.Rows[i].Cells["DENPYONO"].Value.ToString()});
                            if (currentRow != 2)
                            {

                                Excel.Range RngToCopy = worksheet.get_Range("A2").EntireRow;
                                Excel.Range RngToInsert = worksheet.get_Range("A" + (number + 3)).EntireRow;
                                RngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy());
                            }
                            //NO行取得
                            number++;
                            worksheet.Cells[excelRow + 2, 1] = number;

                            for (int j = 0; j < ds.Tables[0].Columns.Count; j++)
                            {
                                worksheet.Cells[excelRow + 2, j + 2] = ds.Tables[0].Rows[excelRow][j];
                            }
                            excelRow++;
                        }
                    }

                }


                catch (Exception e)
                {
                    log.Error("ERROR :" + e.Message);
                }

                finally
                {
                    UpdateUnchin(vs);
                    excel.DisplayAlerts = false;
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

        public void SearchParameter(DataGridView gr)
        {
            try

            {
                DEV10G2U dev = new DEV10G2U();
                StringBuilder sql = new StringBuilder();

                sql.Append("SELECT                                   ");
                sql.Append("T_KDHSINFO.AREACD,                       ");
                sql.Append("T_KDHSINFO.IRIGSYCD,                     ");
                sql.Append("T_KDHSINFO.KURAGO,                   　  ");
                sql.Append("T_KDHSINFO.SYUKABI,                      ");
                sql.Append("T_KDHSINFO.SYUKABI,                      ");
                sql.Append("T_KDHSINFO.NOUKIBI,                      ");
                sql.Append("T_UNCHIN.DHYDENPYONO,                    ");
                sql.Append("T_KDHSINFO.NUKNNM,                       ");
                sql.Append("T_KDHSINFO.ADDRESS,                      ");
                sql.Append("T_UNCHIN.KOSU,                           ");
                sql.Append("T_UNCHIN.WT,                             ");

                //sql.Append("T_KDHSINFO.SOKOCD,                       ");

                sql.Append("T_UNCHIN.SKYUNCHIN,                      ");
                sql.Append("T_UNCHIN.TYUKEIRYO,                      ");
                sql.Append("T_UNCHIN.SNTUNCHINCD1,                   ");
                sql.Append("T_UNCHIN.SNTUNCHINGAK1,                  ");
                sql.Append("T_UNCHIN.SNTUNCHINCD2,                   ");
                sql.Append("T_UNCHIN.SNTUNCHINGAK2,                  ");
                sql.Append("T_UNCHIN.SNTUNCHINCD3,                   ");
                sql.Append("T_UNCHIN.SNTUNCHINGAK3                   ");
                sql.Append("FROM                                     ");
                sql.Append("T_KDHSINFO                               ");
                sql.Append("INNER JOIN  T_UNCHIN                     ");
                sql.Append("ON T_UNCHIN.SOKOCD = T_KDHSINFO.SOKOCD   ");
                sql.Append("AND  T_UNCHIN.DENPYONO = T_KDHSINFO.DENPYONO  ");
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

                sql.Append("    AND T_KDHSINFO.SOKOCD =" + "'" + SOKOCD + "'");
                sql.Append("    AND T_KDHSINFO.ZNKFLG  IS  NULL      ");

                DataSet dataSet = dev.executeSelectQuery(sql.ToString());
                UnchinDownLoad(dataSet, gr);

            }
            catch

            {
                throw;

            }
        }


        //Update database set status to 5
        public void UpdateUnchin(List<string[]> vs)
        {
            DEV10G2U dev = new DEV10G2U();
            foreach (string[] row in vs)
            {
                StringBuilder sql = new StringBuilder();
                sql.Append("UPDATE T_KDHSINFO SET ");
                sql.Append(" STATUS = 5,");
                sql.Append(" LUDATE = to_date('" + DateTime.Now);
                sql.Append("' , 'yyyy-mm-dd hh24:mi:ss'), LUWSID = '" + Environment.MachineName);
                sql.Append("', LUUSERID = '" + Environment.UserName);

                sql.Append("' WHERE ");
                sql.Append(" SOKOCD = '" + row[0] + "' AND");
                sql.Append(" SYKFILENM = '" + row[1] + "' AND");
                sql.Append(" SEQNO = '" + row[2] + "' AND");
                sql.Append(" DENPYONO = '" + row[3] + "'");

                dev.executeUpdateQuery(sql.ToString());
            }
        }
    }

}
