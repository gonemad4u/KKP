using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data.OracleClient;
using System.Xml.Linq;
using System.Windows.Forms;
using System.Diagnostics;

namespace DAL
{
    public class DEV10G2U
    {
        private static readonly log4net.ILog log = log4net.LogManager.GetLogger
               (System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        private OracleConnection conn;
        
        private string connString = "";
        DialogResult myResult;
        //When generated, fill connection string with xml info
        public DEV10G2U()
        {
            var xml = XDocument.Load(@"..\Mitsui.xml");
            var queryC = xml.Root.Descendants("connectionStrings")
                        .Elements("connectionString")
                        .Select(a => a.Value).ToList();
            connString = queryC[0].ToString();
        }

        //Simply executes sql
        public bool executeDeleteQuery(String sql)
        {
            using (conn = new OracleConnection(connString)) {
                conn.Open();
                OracleCommand command = null;
                try
                {
                    command = new OracleCommand(sql, conn);
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    Console.Write("Error - Connection.executeSelectQuery - Query: +  Exception "
                        + e.StackTrace.ToString());
                    return false;
                }
                finally
                {
                    conn.Close();
                }
                return true;
            }
        }

        //Simply executes sql
        public DataTable executeSelectQuery(String sql)
        {
            using (conn = new OracleConnection(connString))
            { 
                DataTable dt = new DataTable();
                try
                {
                    OracleDataAdapter ada = new OracleDataAdapter(sql, conn);
                    ada.Fill(dt);
                }
                catch (Exception e)
                {
                    Console.Write("Error - Connection.executeSelectQuery - Query: +  Exception " 
                        + e.StackTrace.ToString());
                    return null;
                }
                finally
                {
                    conn.Close();
                }
                return dt;
            }
        }
        //Alter the datatable to leave out all the duplicate rows, in case there are no dulicate return false
        public DataTable alterDataTable(DataTable datatable)
        {
            for (int i = datatable.Rows.Count - 1; i >= 0; i--)
            {
                var denpyo = datatable.Rows[i]["DENPYONO"].ToString();
                var irig = datatable.Rows[i]["IRIGSYCD"].ToString();
                var filenm = datatable.Rows[i]["FileName"].ToString();
                var seq = datatable.Rows[i]["SeqNo"].ToString();
                try
                {
                    using (conn = new OracleConnection(connString))
                    {
                        if (checkDuplicate(denpyo, irig, conn))
                        {
                            log.Error("Duplicated!");
                            log.Error("File: " + filenm + ", line: " + seq + ", has duplicates, please check!");
                            datatable.Rows[i].Delete();
                        }
                    }
                }
                catch (Exception e)
                {
                    log.Fatal("Error message: " + e.Message);
                    log.Fatal(e.StackTrace);
                    log.Info("EXEC END");
                    myResult = MessageBox.Show("Unexpected Error!!! See log?", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (myResult == DialogResult.OK)
                    {
                        Process.Start("explorer.exe", @"logs");
                    }
                }
                finally
                {
                    conn.Close();
                }
                
            }
            datatable.AcceptChanges();
            return (datatable);
        }

        //Check the row to see if there is duplicate
        public bool checkDuplicate(string orderNo, string gyosya, OracleConnection conn)
        {
            conn.Open();
            StringBuilder sql = new StringBuilder();
            sql.Append("select count(*) from T_KDHSINFO ");
            List<string> sqlWhere = new List<string>();
            sqlWhere.Add("DENPYONO= '" + orderNo + "'");
            sqlWhere.Add(" IRIGSYCD= '" + gyosya + "'");
            string where = string.Join("  AND   ", sqlWhere.ToArray());
            sql.Append("    WHERE   " + where);
            
            OracleCommand command = new OracleCommand();
            command.CommandText = sql.ToString();
            command.CommandType = CommandType.Text;
            command.Connection = conn;
            object a = command.ExecuteScalar();
            conn.Close();
            int theCount = Convert.ToInt32(a);
            if (theCount == 0)
            {
                return false;
            }
            else
            {
                return true;
            }
        }
        
        //Produce the insert sql from datatable provided, and execute
        public void executeInsertQuery(DataTable datatable)
        {
            int i = 0;
            if (datatable.Rows.Count != 0)
            {
                OracleCommand command = null;
                try
                {
                    using (conn = new OracleConnection(connString))
                    {
                        var commandText = "insert into T_KDHSINFO (SOKOCD,RDDATE,RDTIME,SYKFILENM,SEQNO,STATUS," +
                            "AREACD,DENPYONO,SYUKABI,NOUKIBI,IRIGSYCD," +
                            "NSNNM,NUKNNM,CHIKUCD,POSTCD,ADDRESS,TELNO," +
                            "SKYHINSYUCD,SKYHINSYUNM,KOSU,TANI,WT,SCNDHSTNNM," +
                            "OKURINO,KURAGO,HAISONO,ZNKFLG,UNCHINCALDT,UNCHISNDDT,CDATE,CWSID,CUSRID,BIKO) " +

                            "values(:SOKOCD,:RDDATE,:RDTIME,:SYKFILENM,:SEQNO,:STATUS," +
                            ":AREACD,:DENPYONO,:SYUKABI,:NOUKIBI,:IRIGSYCD," +
                            ":NSNNM,:NUKNNM,:CHIKUCD,:POSTCD,:ADDRESS,:TELNO," +
                            ":SKYHINSYUCD,:SKYHINSYUNM,:KOSU,:TANI,:WT,:SCNDHSTNNM," +
                            ":OKURINO,:KURAGO,:HAISONO,:ZNKFLG,:UNCHINCALDT,:UNCHISNDDT,:CDATE,:CWSID,:CUSRID,:BIKO) ";
                        command = new OracleCommand(commandText, conn);

                        conn.Open();

                        command.Parameters.Add("SOKOCD", OracleType.VarChar);
                        command.Parameters.Add("RDDATE", OracleType.VarChar);
                        command.Parameters.Add("RDTIME", OracleType.VarChar);
                        command.Parameters.Add("SYKFILENM", OracleType.VarChar);
                        command.Parameters.Add("SEQNO", OracleType.Number);
                        command.Parameters.Add("STATUS", OracleType.VarChar);
                        command.Parameters.Add("AREACD", OracleType.VarChar);
                        command.Parameters.Add("DENPYONO", OracleType.VarChar);
                        command.Parameters.Add("SYUKABI", OracleType.VarChar);
                        command.Parameters.Add("NOUKIBI", OracleType.VarChar);
                        command.Parameters.Add("IRIGSYCD", OracleType.VarChar);
                        command.Parameters.Add("NSNNM", OracleType.VarChar);
                        command.Parameters.Add("NUKNNM", OracleType.VarChar);
                        command.Parameters.Add("CHIKUCD", OracleType.VarChar);
                        command.Parameters.Add("POSTCD", OracleType.VarChar);
                        command.Parameters.Add("ADDRESS", OracleType.VarChar);
                        command.Parameters.Add("TELNO", OracleType.VarChar);
                        command.Parameters.Add("SKYHINSYUCD", OracleType.VarChar);
                        command.Parameters.Add("SKYHINSYUNM", OracleType.VarChar);
                        command.Parameters.Add("KOSU", OracleType.Number);
                        command.Parameters.Add("TANI", OracleType.VarChar);
                        command.Parameters.Add("WT", OracleType.Number);
                        command.Parameters.Add("SCNDHSTNNM", OracleType.VarChar);
                        command.Parameters.Add("OKURINO", OracleType.VarChar);
                        command.Parameters.Add("KURAGO", OracleType.VarChar);
                        command.Parameters.Add("HAISONO", OracleType.VarChar);
                        command.Parameters.Add("ZNKFLG", OracleType.VarChar);
                        command.Parameters.Add("UNCHINCALDT", OracleType.VarChar);
                        command.Parameters.Add("UNCHISNDDT", OracleType.VarChar);
                        command.Parameters.Add("CDATE", OracleType.DateTime);
                        command.Parameters.Add("CWSID", OracleType.VarChar);
                        command.Parameters.Add("CUSRID", OracleType.VarChar);
                        command.Parameters.Add("CUSRID", OracleType.VarChar);
                        command.Parameters.Add("BIKO", OracleType.VarChar);

                        foreach (DataRow row in datatable.Rows)
                        {
                            command.Parameters["SOKOCD"].Value = Tools.getSokocd();
                            command.Parameters["RDDATE"].Value = DateTime.Now.ToString("yyyy/MM/dd");
                            command.Parameters["RDTIME"].Value = DateTime.Now.ToString("HHmmss");
                            command.Parameters["SYKFILENM"].Value = row["FileName"].ToString();
                            command.Parameters["SEQNO"].Value = row["SeqNo"].ToString();
                            command.Parameters["STATUS"].Value = 0;
                            command.Parameters["AREACD"].Value = row["AREACD"].ToString();
                            command.Parameters["DENPYONO"].Value = row["DENPYONO"].ToString();
                            command.Parameters["SYUKABI"].Value = row["SYUKABI"].ToString();
                            command.Parameters["NOUKIBI"].Value = row["NOUKIBI"].ToString();
                            command.Parameters["IRIGSYCD"].Value = row["IRIGSYCD"];
                            command.Parameters["NSNNM"].Value = DBNull.Value;
                            command.Parameters["NUKNNM"].Value = row["NUKNNM"].ToString();
                            command.Parameters["CHIKUCD"].Value = row["CHIKUCD"].ToString();
                            command.Parameters["POSTCD"].Value = DBNull.Value;
                            command.Parameters["ADDRESS"].Value = row["ADDRESS"].ToString();
                            command.Parameters["TELNO"].Value = row["TELNO"].ToString();
                            command.Parameters["SKYHINSYUCD"].Value = row["SKYHINSYUCD"].ToString();
                            command.Parameters["SKYHINSYUNM"].Value = row["SKYHINSYUNM"].ToString();
                            command.Parameters["KOSU"].Value = row["KOSU"].ToString();
                            command.Parameters["TANI"].Value = DBNull.Value;
                            command.Parameters["WT"].Value = row["WT"].ToString();
                            command.Parameters["SCNDHSTNNM"].Value = DBNull.Value;
                            command.Parameters["OKURINO"].Value = DBNull.Value;
                            command.Parameters["KURAGO"].Value = row["KURAGO"].ToString();
                            command.Parameters["HAISONO"].Value = DBNull.Value;
                            command.Parameters["ZNKFLG"].Value = DBNull.Value;
                            command.Parameters["UNCHINCALDT"].Value = DBNull.Value;
                            command.Parameters["UNCHISNDDT"].Value = DBNull.Value;
                            command.Parameters["CDATE"].Value = DateTime.Now;
                            command.Parameters["CWSID"].Value = Environment.MachineName;
                            command.Parameters["CUSRID"].Value = Environment.UserName;
                            command.Parameters["BIKO"].Value = row["MEMO"].ToString();
                            command.ExecuteNonQuery();
                            i++;
                        }
                    }
                }
                catch (Exception e)
                {
                    log.Fatal("Error message: " + e.Message);
                    log.Fatal(e.StackTrace);
                    log.Info("EXEC END");
                    myResult = MessageBox.Show("Unexpected Error!!! See log?", "Error", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);
                    if (myResult == DialogResult.OK)
                    {
                        Process.Start("explorer.exe", @"logs");
                    }
                }
                finally
                {
                    log.Info("Successfully inserted " + i + "records.");
                    conn.Close();
                }
            }
        }

        //Alternative search method, to search from datable, used in loading Syuyaku only
        public DataTable searchBy(DataTable dt)
        {
            if(dt.Rows.Count != 0)
            {
                StringBuilder sql = new StringBuilder();
                sql.Append("SELECT                                   ");
                sql.Append("t.*,    ");
                sql.Append("CODE_STATUS.CODE1,        ");
                sql.Append("CODE_SOKO.CODENAME,        ");
                sql.Append("Gyosya.SOKONM        ");
                sql.Append("FROM                                     ");
                sql.Append("T_KDHSINFO  t                             ");
                sql.Append("INNER JOIN  M_MNRCODE  CODE_STATUS          ");
                sql.Append("ON   CODE_STATUS.CODE = t.STATUS   ");
                sql.Append("AND  CODE_STATUS.SOKOCD = 'DEF'             ");
                sql.Append("AND  CODE_STATUS.BNRICODE = 'STATUS'        ");
                sql.Append(" INNER JOIN  M_MNRCODE  CODE_SOKO           ");
                sql.Append("ON   CODE_SOKO.CODE = t.IRIGSYCD   ");
                sql.Append("AND  CODE_SOKO.SOKOCD = 'DEF'               ");
                sql.Append("AND  CODE_SOKO.BNRICODE = 'SOKOCD'          ");
                sql.Append(" INNER JOIN  M_SOKO  Gyosya           ");
                sql.Append("ON   Gyosya.SOKOCD = t.IRIGSYCD   ");
                sql.Append("    WHERE                                ");
                string denpyonoSql = "t.DENPYONO IN(      ";
                string sykFileNMSql = "AND t.SYKFILENM IN (     ";
                string seqNoSql = "AND t.SEQNO IN (     ";

                foreach (DataRow row in dt.Rows)
                {
                    denpyonoSql += "'" + row["DENPYONO"] + "'" + ",";
                    sykFileNMSql += "'" + row["FileName"] + "'" + ",";
                    seqNoSql += "'" + row["SeqNo"] + "'" + ",";
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
                return executeSelectQuery(sql.ToString());
            }
            else
            {
                return new DataTable();
            }
        }

        //Simply executes sql
        public bool executeUpdateQuery(String sql)
        {
            using (conn = new OracleConnection(connString))
            {
                conn.Open();
                OracleCommand command = null;
                try
                {
                    command = new OracleCommand(sql, conn);
                    command.ExecuteNonQuery();
                }
                catch (Exception e)
                {
                    Console.Write("Error - Connection.executeSelectQuery - Query: +  Exception "
                        + e.StackTrace.ToString());
                    return false;
                }
                finally
                {
                    conn.Close();
                }
                return true;
            }
        }

    }
}
