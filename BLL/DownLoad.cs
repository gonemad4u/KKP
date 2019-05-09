using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;
using DAL;
using System.Runtime.InteropServices;
using System.Xml.Linq;
using System.Linq;

namespace BLL
{
    public class DownLoad
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
              (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void SyuYakuDL(DataGridView gr, DataTable ds)
        {
            
            string saveFileName = "集約_" + DateTime.Now.ToString("yyyyMMdd");
            string templetFile = @"template\集約.xlsx";
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
                worksheet.Name = "集約";
                excel.Visible = false ;

                int excelRow = 0;
                int number = 1;
                int selectRow = 0;
                
                try
                {
                    log.Info("EXEC BEGIN");
                    //一行目値取得
                    //datagridview行数判断
                    for (int a = 0; a < gr.RowCount; a++)
                    {
                        //選択行
                        if ((bool)gr.Rows[a].Cells[0].EditedFormattedValue == true)
                        {
                            //DataSet行数判断
                            for (int i = 0; i < ds.Rows.Count; i++)
                            {
                                if (gr.Rows[a].Cells["DENPYONO"].Value.Equals(ds.Rows[i][13])
                                 && gr.Rows[a].Cells["SOKOCD"].Value.Equals(ds.Rows[i][19])
                                 && gr.Rows[a].Cells["SYKFILENM"].Value.Equals(ds.Rows[i][17])
                                 && gr.Rows[a].Cells["SEQNO"].Value.Equals(ds.Rows[i][18]))
                                {
                                    //NO行取得
                                    worksheet.Cells[2, 1] = 1;
                                    for (int r = 0; r < 19; r++)
                                    {
                                        worksheet.Cells[2, r + 2] = ds.Rows[i][r];                          
                                    }
                                    selectRow = a;
                                    break;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            break;
                        }
                    }

                    for (int b = selectRow + 1; b < gr.RowCount; b++)
                    {
                        if ((bool)gr.Rows[b].Cells[0].EditedFormattedValue == true)
                        {
                            for (int i = 0; i < ds.Rows.Count; i++)
                            {
                                 if (gr.Rows[b].Cells["DENPYONO"].Value.Equals(ds.Rows[i][13])
                                    && gr.Rows[b].Cells["SOKOCD"].Value.Equals(ds.Rows[i][19])
                                    && gr.Rows[b].Cells["SYKFILENM"].Value.Equals(ds.Rows[i][17])
                                    && gr.Rows[b].Cells["SEQNO"].Value.Equals(ds.Rows[i][18]))
                                 {                                    
                                    Excel.Range RngToCopy = worksheet.get_Range("A2").EntireRow;
                                    Excel.Range RngToInsert = worksheet.get_Range("A" + (number + 2)).EntireRow;
                                    RngToInsert.Insert(Excel.XlInsertShiftDirection.xlShiftDown, RngToCopy.Copy());
                   
                                    //NO行取得
                                    number++;
                                    worksheet.Cells[excelRow + 3, 1] = number;

                                    for (int j = 0; j < 19; j++)
                                    {
                                        worksheet.Cells[excelRow + 3, j + 2] = ds.Rows[i][j];
                                    }
                                        excelRow++;
                                 }
                            }
                        }
                    }
                }

                catch (Exception e)
                {
                    log.Error("ERROR" + e.Message);
                }

                finally
                {
                    Application.DoEvents();
                    var xml = XDocument.Load(@"..\Mitsui.xml");
                    var queryC = xml.Root.Descendants("rndflg")
                                .Elements("col")
                                .Select(a => a.Value);

                    List<string> cols = new List<string>();

                        
                    foreach (var element in queryC)
                    {
                        cols = element.Split(',').ToList();
                    }
                    
                    foreach (var b in cols)
                    {            
                        worksheet.get_Range(b + ":" + b).Locked = false;
                    }
                    worksheet.Protect();

                    //workbook.Saved = true;
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

        public DataTable SearchParameter(bool MNR, bool SNK, bool TYU, bool CYU, bool APU, bool Status1, bool Status2, bool Status3, bool Status4, bool Status5, bool Status6, bool Status7,
            string SyukaBi1, string SyukaBi2, string Exlsrd1, string Exlsrd2, string Unchin1, string Unchin2,
            string OrderNo, string HaisoBnNO, bool Zanka, object Area, bool SyukaFlgL, bool SyukaFlgR, bool ExlFlgL, bool ExlFlgR, bool UnchinFlgL, bool UnchinFlgR)
        {
            try

            {
                DEV10G2U dev = new DEV10G2U();
                StringBuilder sql = new StringBuilder();

                sql.Append("SELECT                    ");
                sql.Append("T_KDHSINFO.AREACD,        ");
                sql.Append("T_KDHSINFO.CHIKUCD,       ");
                sql.Append("T_KDHSINFO.NUKNNM,        ");
                sql.Append("T_KDHSINFO.SYUKABI,       ");
                sql.Append("T_KDHSINFO.NOUKIBI,       ");
                sql.Append("T_KDHSINFO.KOSU,          ");
                sql.Append("T_KDHSINFO.WT,            ");
                sql.Append("T_KDHSINFO.HAISONO,       ");
                sql.Append("T_KDHSINFO.ZNKFLG,        ");
                sql.Append("T_KDHSINFO.IRIGSYCD,      ");
                //sql.Append("M_SOKO.SOKONM,            ");
                sql.Append("T_KDHSINFO.KURAGO,        ");
                sql.Append("T_KDHSINFO.SKYHINSYUCD,   ");
                sql.Append("T_KDHSINFO.SKYHINSYUNM,   ");
                sql.Append("T_KDHSINFO.DENPYONO,      ");
                sql.Append("T_KDHSINFO.ADDRESS,       ");
                sql.Append("T_KDHSINFO.TELNO,         ");
                sql.Append("T_KDHSINFO.BIKO,          ");

                sql.Append("T_KDHSINFO.SYKFILENM,     ");
                sql.Append("T_KDHSINFO.SEQNO,         ");
                sql.Append("T_KDHSINFO.SOKOCD         ");

                sql.Append("FROM                      ");
                sql.Append("T_KDHSINFO                ");
                //sql.Append("INNER JOIN  M_SOKO        ");
                //sql.Append("ON   M_SOKO.SOKOCD = T_KDHSINFO.IRIGSYCD   ");

                if (Area != null)
                {
                    sql.Append(" ,(SELECT              ");
                    sql.Append("AREANM ,  AREACD       ");
                    sql.Append("FROM                   ");
                    sql.Append("M_AREA                 ");
                    sql.Append("GROUP BY AREANM, AREACD ");
                    sql.Append(")AREA                  ");

                }
                List<string> sqlWhere = new List<string>();
                if (Area != null)
                {
                    sqlWhere.Add("    AREA.AREACD = T_KDHSINFO.AREACD" + "  AND  AREA.AREANM= '" + Area.ToString() + "'");
                }
                var companySql = " CODE_SOKO.KBN1 IN( ";
                bool companySelected = false;
                if (MNR)
                {
                    companySql += "1,";
                    companySelected = true;
                }
                if (SNK)
                {
                    companySql += "2,";
                    companySelected = true;
                }
                if (TYU)
                {
                    companySql += "3,";
                    companySelected = true;
                }
                if (CYU)
                {
                    companySql += "4,";
                    companySelected = true;
                }
                if (APU)
                {
                    companySql += "5,";
                    companySelected = true;
                }
                if (companySelected)
                {
                    companySql = companySql.Remove(companySql.Length - 1);
                    sqlWhere.Add(companySql + ")");
                }

                //check if unchin has been sent
                bool unchinFlg = false;
                if (Status1)
                {
                    sqlWhere.Add("STATUS=" + '0');
                }

                if (Status2)
                {
                    sqlWhere.Add("STATUS=" + '2');
                }

                if (Status3)
                {
                    sqlWhere.Add("STATUS=" + '1');
                }

                if (Status4)
                {
                    sqlWhere.Add("STATUS=" + '3');
                }

                if (Status5)
                {
                    sqlWhere.Add("STATUS=" + '5');
                }

                if (Status7)
                {
                    unchinFlg = true;
                    sqlWhere.Add("STATUS=" + '4');
                }

                if (Status6)
                {
                    unchinFlg = true;
                }

                if (SyukaBi1.Length > 0 && SyukaFlgL)
                {
                    sqlWhere.Add("SYUKABI>= " + "'" + SyukaBi1.Substring(0, 10) + "'");
                }

                if (SyukaBi2.Length > 0 && SyukaFlgR)
                {
                    sqlWhere.Add("SYUKABI <=" + "'" + SyukaBi2.Substring(0, 10) + "'");
                }

                if (Exlsrd1.Length > 0 && ExlFlgL)
                {
                    sqlWhere.Add("RDDATE>= " + "'" + Exlsrd1.Substring(0, 10) + "'");
                }

                if (Exlsrd2.Length > 0 && ExlFlgR)
                {
                    sqlWhere.Add("RDDATE<= " + "'" + Exlsrd2.Substring(0, 10) + "'");
                }

                if (unchinFlg)
                { 
                    if (Unchin1.Length > 0 && UnchinFlgL)
                    {
                        sqlWhere.Add("UNCHISNDDT>= " + "'" + Unchin1.Substring(0, 10) + "'");
                    }

                    if (Unchin2.Length > 0 && UnchinFlgR)
                    {
                        sqlWhere.Add("UNCHISNDDT < = " + "'" + Unchin2.Substring(0, 10) + "'");
                    }
                }
                if (OrderNo.Length > 0)
                {
                    sqlWhere.Add("DENPYONO= '" + OrderNo + "'");
                }

                if (HaisoBnNO.Length > 0)
                {
                    sqlWhere.Add("HAISONO= '" + HaisoBnNO + "'");
                }

                if (Zanka)
                {
                    sqlWhere.Add("ZNKFLG=  'Y' ");
                }
                else
                {
                    //sqlWhere.Add("ZNKFLG IS NULL ");
                }

                if (sqlWhere.Count > 0)
                {
                    string count = string.Join("  AND   ", sqlWhere.ToArray());
                    sql.Append("    WHERE   " + count);
                }
                return dev.executeSelectQuery(sql.ToString());
            }

            catch 
            {

                throw;

            }
        }


    }
}