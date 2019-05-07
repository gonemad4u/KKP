using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DAL;
using System.Xml.Linq;

namespace BLL
{
    public class HaisoDl
    {
        public string CreateExcel(DataGridView gr)
        {
            string time = DateTime.Now.ToString("yyyyMMdd");
            string saveFileName = "配車依頼表フォーム_" + time + ".xlsx";
            SaveFileDialog saveDialog = new SaveFileDialog
            {
                Filter = "Excel文件 |*.xlsx",
                FileName = saveFileName,
                RestoreDirectory = true
            };

            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                saveFileName = saveDialog.FileName;
                File.Copy(@"template\配車依頼表フォーム_yyyymmdd.xlsx", saveFileName, true);
                var excelApp = new Excel.Application();
                var workbook = excelApp.Workbooks.Open(saveFileName);
                var sheet = (Excel.Worksheet)workbook.Worksheets[1];
                excelApp.Visible = false;
                int currentRow = 2;
                List<string[]> vs = new List<string[]>();//Create a sring list that stores the four keys
                DataTable dt = searchBy(gr);
                try
                {

                    for (int i = 0; i < gr.RowCount; i++)
                    {
                        if ((bool)gr.Rows[i].Cells[0].EditedFormattedValue == true)
                        {
                            //First we extract the keys of rows selected
                            vs.Add(new string[]{ gr.Rows[i].Cells["SOKOCD"].Value.ToString(),
                        gr.Rows[i].Cells["SYKFILENM"].Value.ToString(),
                        gr.Rows[i].Cells["SEQNO"].Value.ToString(),
                        gr.Rows[i].Cells["DENPYONO"].Value.ToString()});
                            //Then we fill the excel file with data
                            //if it's the first row we only fill in data, else we create a new row

                            if (currentRow != 2)
                            {
                                Excel.Range from = sheet.get_Range("B" + (currentRow - 1)).EntireRow;
                                Excel.Range to = sheet.get_Range("B" + currentRow).EntireRow;
                                to.Insert(Excel.XlInsertShiftDirection.xlShiftDown, from.Copy());
                            }
                            for (int j = 0; j < dt.Columns.Count; j++)
                            {
                                sheet.Cells[currentRow, j + 1] = dt.Rows[currentRow - 2][j].ToString();
                            }
                            currentRow++;
                        }
                    }
                }
                catch (Exception e)
                {
                    return "nm$l:" + e.Message;
                }
                finally
                {
                    UpdateByKey(vs);
                    excelApp.DisplayAlerts = false;
                    workbook.Save();
                    workbook.Close(false);
                    Marshal.FinalReleaseComObject(excelApp);
                }
                return "niubi";
            }
            else
            {
                MessageBox.Show("処理を中止しました。");
            
                return "cancel";
            }
        }
        public DataTable searchBy(DataGridView gr)
        {
            StringBuilder sql = new StringBuilder();
            sql.Append("SELECT                                   ");
            sql.Append("t.DENPYONO,                       ");
            sql.Append("t.SYUKABI,                     ");
            sql.Append("t.NOUKIBI,                      ");
            sql.Append("t.IRIGSYCD,                     ");
            sql.Append("m.SOKONM,                            ");
            sql.Append("t.NSNNM,                       ");
            sql.Append("t.NUKNNM,                      ");
            sql.Append("t.CHIKUCD,                         ");
            sql.Append("t.POSTCD,                           ");
            sql.Append("t.ADDRESS,                         ");
            sql.Append("t.TELNO,                         ");
            sql.Append("t.KOSU,                         ");
            sql.Append("t.TANI,                         ");
            sql.Append("t.WT,                         ");
            sql.Append("t.SCNDHSTNNM,                         ");
            sql.Append("t.OKURINO,                         ");
            sql.Append("t.BIKO,                         ");
            sql.Append("t.KURAGO                         ");
            sql.Append("FROM                                     ");
            sql.Append("T_KDHSINFO  t                             ");
            sql.Append("INNER JOIN  M_SOKO m                      ");
            sql.Append("ON   m.SOKOCD = t.IRIGSYCD ");
            sql.Append("    WHERE                                ");
            string denpyonoSql = "t.DENPYONO IN(      ";
            string sykFileNMSql = "AND t.SYKFILENM IN (     ";
            string seqNoSql = "AND t.SEQNO IN (     ";

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
            sql.Append(" AND t.SOKOCD = '");
            
            var SOKOCD = Tools.getSokocd();
            sql.Append(SOKOCD + "'");
            sql.Append(" AND t.STATUS = 2");
            DEV10G2U dev = new DEV10G2U();
            return dev.executeSelectQuery(sql.ToString());
        }
        public void UpdateByKey(List<string[]> vs)
        {
            DEV10G2U dev = new DEV10G2U();
            foreach (string[] row in vs)
            {
                StringBuilder sql = new StringBuilder();
                sql.Append("UPDATE T_KDHSINFO SET ");
                sql.Append(" STATUS = 3,");
                sql.Append(" LUDATE = to_date('" + DateTime.Now);
                sql.Append("' , 'yyyy-mm-dd hh24:mi:ss'), LUWSID = '" + Environment.MachineName);
                sql.Append("', LUUSERID = '" + Environment.UserName);

                sql.Append("' WHERE ");
                sql.Append(" SOKOCD = '" + row[0] + "' AND");
                sql.Append(" SYKFILENM = '" + row[1] + "' AND");
                sql.Append(" SEQNO = '" + row[2] + "' AND");
                sql.Append(" DENPYONO = '" + row[3] + "' AND");
                sql.Append(" STATUS = 2");

                dev.executeUpdateQuery(sql.ToString());
            }
        }
        
    }
}
