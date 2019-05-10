using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using DAL;
using System.Xml.Linq;

namespace BLL
{
    public class Search
    {
        public DataTable SearchByParameter(bool MNR, bool SNK, bool TYU, bool CYU, bool APU, bool Status1, bool Status2, bool Status3, bool Status4, bool Status5, bool Status6, bool Status7,
            string SyukaBi1, string SyukaBi2, string Exlsrd1, string Exlsrd2, string Unchin1, string Unchin2,
            string OrderNo, string HaisoBnNO, bool Zanka, object Area, bool SyukaFlgL, bool SyukaFlgR, bool ExlFlgL, bool ExlFlgR, bool UnchinFlgL, bool UnchinFlgR)
        {
            try

            {
                DEV10G2U dal = new DEV10G2U();
                StringBuilder sql = new StringBuilder();
                sql.Append("SELECT                    ");
                sql.Append("t.SOKOCD,        ");
                sql.Append("t.RDDATE,        ");
                sql.Append("t.RDTIME,        ");
                sql.Append("t.SYKFILENM,     ");
                sql.Append("t.SEQNO,         ");
                sql.Append("t.STATUS,        ");
                sql.Append("t.AREACD,        ");
                sql.Append("t.DENPYONO,      ");
                sql.Append("t.SYUKABI,       ");
                sql.Append("t.NOUKIBI,       ");
                sql.Append("t.IRIGSYCD,      ");
                //sql.Append("T_KDHSINFO.IRIGSYNM,      ");
                sql.Append("t.NSNNM,         ");
                sql.Append("t.NUKNNM,        ");
                sql.Append("t.CHIKUCD,       ");
                sql.Append("t.POSTCD,        ");
                sql.Append("t.ADDRESS,       ");
                sql.Append("t.TELNO,         ");
                sql.Append("t.SKYHINSYUCD,   ");
                sql.Append("t.SKYHINSYUNM,   ");
                sql.Append("t.KOSU ,         ");
                sql.Append("t.TANI,          ");
                sql.Append("t.WT,            ");
                sql.Append("t.SCNDHSTNNM,    ");
                sql.Append("t.OKURINO,       ");
                sql.Append("t.KURAGO,        ");
                sql.Append("t.HAISONO,       ");
                sql.Append("t.ZNKFLG,        ");
                sql.Append("t.UNCHINCALDT,   ");
                sql.Append("t.UNCHISNDDT,    ");
                sql.Append("Gyosya.SOKONM,        ");
                sql.Append("CODE_STATUS.CODE1,        ");
                sql.Append("CODE_SOKO.CODENAME        ");
                sql.Append("FROM                      ");
                sql.Append("T_KDHSINFO t                ");
                sql.Append(" INNER JOIN  M_SOKO  Gyosya           ");
                sql.Append("ON   Gyosya.SOKOCD = t.IRIGSYCD   ");
                sql.Append("INNER JOIN  M_MNRCODE  CODE_STATUS          ");
                sql.Append("ON   CODE_STATUS.CODE = t.STATUS   ");
                sql.Append("AND  CODE_STATUS.SOKOCD = 'DEF'             ");
                sql.Append("AND  CODE_STATUS.BNRICODE = 'STATUS'        ");
                sql.Append(" INNER JOIN  M_MNRCODE  CODE_SOKO           ");
                sql.Append("ON   CODE_SOKO.CODE = t.IRIGSYCD   ");
                sql.Append("AND  CODE_SOKO.SOKOCD = 'DEF'               ");
                sql.Append("AND  CODE_SOKO.BNRICODE = 'SOKOCD'          ");

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
                    sqlWhere.Add("    AREA.AREACD = t.AREACD" + "  AND  AREA.AREANM = '" + Area.ToString() + "'");
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
                    companySql += ")";
                    sqlWhere.Add(companySql);
                }

                //check if unchin has been sent
                bool unchinFlg = false;
                if (Status1)
                {
                    sqlWhere.Add("t.STATUS=" + '0');
                }

                if (Status2)
                {
                    sqlWhere.Add("t.STATUS=" + '1');
                }

                if (Status3)
                {
                    sqlWhere.Add("t.STATUS=" + '2');
                }

                if (Status4)
                {
                    sqlWhere.Add("t.STATUS=" + '3');
                }

                if (Status5)
                {
                    sqlWhere.Add("t.STATUS=" + '5');
                }

                if (Status7)
                {
                    unchinFlg = true;
                    sqlWhere.Add("t.STATUS=" + '4');
                }

                if (Status6)
                {
                    unchinFlg = true;
                }

                if (SyukaBi1.Length > 0 && SyukaFlgL)
                {
                    sqlWhere.Add("t.SYUKABI>= " + "'" + SyukaBi1.Substring(0, 10) + "'");
                }

                if (SyukaBi2.Length > 0 && SyukaFlgR)
                {
                    sqlWhere.Add("t.SYUKABI <=" + "'" + SyukaBi2.Substring(0, 10) + "'");
                }

                if (Exlsrd1.Length > 0 && ExlFlgL)
                {
                    sqlWhere.Add("t.RDDATE>= " + "'" + Exlsrd1.Substring(0, 10) + "'");
                }

                if (Exlsrd2.Length > 0 && ExlFlgR)
                {
                    sqlWhere.Add("t.RDDATE<= " + "'" + Exlsrd2.Substring(0, 10) + "'");
                }

                if (unchinFlg)
                {
                    if (Unchin1.Length > 0 && UnchinFlgL)
                    {
                        sqlWhere.Add("t.UNCHISNDDT>= " + "'" + Unchin1.Substring(0, 10) + "'");
                    }

                    if (Unchin2.Length > 0 && UnchinFlgR)
                    {
                        sqlWhere.Add("t.UNCHISNDDT < = " + "'" + Unchin2.Substring(0, 10) + "'");
                    }
                }

                if (OrderNo.Length > 0)
                {
                    sqlWhere.Add("t.DENPYONO= '" + OrderNo + "'");
                }

                if (HaisoBnNO.Length > 0)
                {
                    sqlWhere.Add("t.HAISONO= '" + HaisoBnNO + "'");
                }

                if (Zanka)
                {
                    sqlWhere.Add("t.ZNKFLG=  'Y' ");
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
                var SOKOCD = Tools.getSokocd();

                sql.Append("    AND t.SOKOCD =" + "'" + SOKOCD + "'");

                return dal.executeSelectQuery(sql.ToString());

            }

            catch

            {

                throw;

            }
        }
    }
}