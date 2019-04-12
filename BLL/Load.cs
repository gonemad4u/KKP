using System;
using System.IO;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Diagnostics;
using DAL;

namespace BLL
{
    public class Load
    {
        private bool FirstFileFlag = true;
        private DataTable dt = new DataTable();
        DataTable dtAlter = new DataTable();
        DEV10G2U dev = new DEV10G2U();
        DialogResult myResult;
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
              (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        //Changes the column header name from Japanese to English
        public string change(string columnName)
        {
            Dictionary<string, string> Data_Array = new Dictionary<string, string>();
            Data_Array.Add("エリアコード", "AREACD");
            Data_Array.Add("指図№", "DENPYONO");
            Data_Array.Add("受付日", "SYUKABI");
            Data_Array.Add("納期", "NOUKIBI");
            Data_Array.Add("倉庫会社コード", "IRIGSYCD");
            Data_Array.Add("荷受人名", "NUKNNM");
            Data_Array.Add("荷受人地域コード", "CHIKUCD");
            Data_Array.Add("荷受人住所", "ADDRESS");
            Data_Array.Add("荷受人電話番号", "TELNO");
            Data_Array.Add("請求品種コード", "SKYHINSYUCD");
            Data_Array.Add("請求品種名", "SKYHINSYUNM");
            Data_Array.Add("個数", "KOSU");
            Data_Array.Add("重量", "WT");
            Data_Array.Add("倉庫コード", "KURAGO");
            Data_Array.Add("送信時間", "SENDTIME");
            Data_Array.Add("記事", "MEMO");
            return Data_Array[columnName];
        }

        //Read a single csv file and transforms it into a datatable
        public void csvToDT(FileInfo f)
        {
            try { 
                int line = 1;
                if (FirstFileFlag)
                {
                    using (StreamReader sr = new StreamReader(f.FullName, Encoding.GetEncoding("Shift_JIS")))
                    {
                        string[] headers = sr.ReadLine().Split(',');

                        foreach (string header in headers)
                        {
                            dt.Columns.Add(change(header));
                        }
                        dt.Columns.Add("FileName");
                        dt.Columns.Add("SeqNo");
                        while (!sr.EndOfStream)
                        {
                        
                            string[] rows = sr.ReadLine().Split(',');
                            DataRow dr = dt.NewRow();
                            for (int i = 0; i < headers.Length - 1; i++)
                            {
                                dr[i] = rows[i];
                            }
                            dr[headers.Length] = f.Name;
                            dr[headers.Length + 1] = line++;
                            dt.Rows.Add(dr);
                        }
                    }
                    line = 0;
                    FirstFileFlag = false;
                }
                else
                {
                    using (StreamReader sr = new StreamReader(f.FullName, Encoding.GetEncoding("Shift_JIS")))
                    {
                        sr.ReadLine();
                        while (!sr.EndOfStream)
                        {
                            string[] rows = sr.ReadLine().Split(',');
                            DataRow dr = dt.NewRow();
                            for (int i = 0; i < dt.Columns.Count - 2; i++)
                            {
                                dr[i] = rows[i];
                            }
                            dr[dt.Columns.Count - 2] = f.Name;
                            dr[dt.Columns.Count - 1] = line++; ;
                            dt.Rows.Add(dr);
                        }
                        line = 0;
                    }
                }
                //Copy file to bak folder
                string targetPath = f.DirectoryName + "\\bak";
                if (!Directory.Exists(targetPath))
                {
                    Directory.CreateDirectory(targetPath);
                }
                File.Move(f.FullName, targetPath + "\\" + f.Name);
            }
            catch(Exception e) { 
                log.Fatal("Error message: " + e.Message);
                log.Fatal(e.StackTrace);
                log.Info("EXEC END");
                myResult = MessageBox.Show("CSVファイルのレイアウトをチェックしてください。ログを見ますか？", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (myResult == DialogResult.OK)
                {
                    Process.Start("explorer.exe", @"logs");
                }
            }
        }
        
        //Insert datatable created from csv, into database
        public void insertData()
        {
            int a = dt.Rows.Count;
            dtAlter = dev.alterDataTable(dt);
            dev.executeInsertQuery(dtAlter);
            if (dtAlter.Rows.Count != a)
            {
                
                myResult = MessageBox.Show("重複データあり、ログを見ますか？", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                if (myResult == DialogResult.OK)
                {
                    Process.Start("explorer.exe", @"logs");
                }
            }
        }

        //Provide the interface with datatable
        public DataTable getData()
        {
            DataTable searched = dev.searchBy(dtAlter);
            return searched;
        }
    }
}
