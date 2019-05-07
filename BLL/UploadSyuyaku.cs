using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.OleDb;
using System.Data;
using DAL;
using System.Xml.Linq;

namespace BLL
{
    public class UploadSyuyaku
    {
        //Read info from Excel and return a datatable
        public DataTable ReadExcel(string fullPath)
        {
            string path = @fullPath;
            string sheetName = "集約";
            using (OleDbConnection conn = new OleDbConnection())
            {
                DataTable dt = new DataTable();
                string Import_FileName = path;
                conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";
                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + sheetName + "$]";

                    comm.Connection = conn;

                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                        return dt;
                    }
                }
            }
        }
        //Update to database
        public void UpdateByDatable(string path)
        {
            DataTable dt = ReadExcel(path);
            DEV10G2U dev = new DEV10G2U();
            
            var xml = XDocument.Load(@"..\Mitsui.xml");
            var queryC = xml.Root.Descendants("rndflg")
                        .Elements("col")
                        .Select(a => a.Value);
            var queryS = xml.Root.Descendants("rndflg")
                        .Elements("set")
                        .Select(a => a.Value);
            string SOKOCD = Tools.getSokocd();
            List<string> cols = new List<string>();
            List<string> sets = new List<string>();

            foreach (var element in queryC)
            {
                cols = element.Split(',').ToList();
            }
            foreach (var element in queryS)
            {
                sets = element.Split(',').ToList();
            }

            
            foreach (DataRow rw in dt.Rows)
            {
                StringBuilder sql = new StringBuilder();
                sql.Append("UPDATE T_KDHSINFO SET ");
                for (int i = 0; i < cols.Count; i++)
                {
                    var col = cols[i];
                    var set = sets[i];
                    var colVal = rw[((int)char.Parse(col) % 32) - 1];
                    string ZNK = rw[((int)'J' % 32) - 1].ToString();
                    string HAISO = rw[((int)'I' % 32) - 1].ToString();
                    if (ZNK == "Y" && HAISO != "")
                    {
                        throw new ArgumentException("NMSL"); ;
                    }
                    sql.Append(set + " = '"+ colVal + "',");
                }
                sql.Append(" STATUS = 1,");
                sql.Append(" LUDATE = to_date('" + DateTime.Now);
                sql.Append("' , 'yyyy-mm-dd hh24:mi:ss'), LUWSID = '" + Environment.MachineName);
                sql.Append("', LUUSERID = '" + Environment.UserName);
                var SYKFILENM = rw[((int) 'S' % 32) - 1];
                var SEQNO = rw[((int)'T' % 32) - 1];
                var DENPYONO = rw[((int)'O' % 32) - 1];
                //
               
                sql.Append("' WHERE ");
                sql.Append(" SOKOCD = '" + SOKOCD + "' AND");
                sql.Append(" SYKFILENM = '" + SYKFILENM + "' AND");
                sql.Append(" SEQNO = '" + SEQNO + "' AND");
                sql.Append(" DENPYONO = '" + DENPYONO + "'");
                //sql.Append(" STATUS = '1'");
                //
                dev.executeUpdateQuery(sql.ToString());
            }
        }
    }
}
