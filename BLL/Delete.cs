using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace BLL
{
    public class Delete
    {
        //Execute the delete proccess from the datagridview
        public bool DeleteFromGL(DataGridView gr)
        {
            List<string[]> vs = new List<string[]>();//Create a sring list that stores the four keys
            for (int i = 0; i < gr.RowCount; i++)
            {
                if ((bool)gr.Rows[i].Cells[0].EditedFormattedValue == true)
                {
                    vs.Add(new string[]{ gr.Rows[i].Cells["DENPYONO"].Value.ToString(),
                    gr.Rows[i].Cells["SOKOCD"].Value.ToString(),
                    gr.Rows[i].Cells["SYKFILENM"].Value.ToString(),
                    gr.Rows[i].Cells["SEQNO"].Value.ToString()});
                }
            }
            var a = CreateDeleteSql(vs); //Create the delete sql from the list of keys provided
            DAL.DEV10G2U d = new DAL.DEV10G2U();
            d.executeDeleteQuery(a);//execute the delete sql
            return true;
        }

        //Create the delete sql, note that the number of keys is 4, so the combinition won't duplicate
        public string CreateDeleteSql(List<string[]> str)
        {
            StringBuilder sql = new StringBuilder();
            sql.Append("DELETE FROM T_KDHSINFO WHERE ");
            sql.Append(" DENPYONO in (");
            for (int i = 0; i < str.Count; i++)
            {
                sql.Append("'" + str[i][0] + "'" + ",");
            }
            sql.Length--;
            sql.Append(") AND SOKOCD in (");
            for (int i = 0; i < str.Count; i++)
            {
                sql.Append("'" + str[i][1] + "'"  + ",");
            }
            sql.Length--;
            sql.Append(") AND SYKFILENM in (");
            for (int i = 0; i < str.Count; i++)
            {
                sql.Append("'" + str[i][2] + "'"  + ",");
            }
            sql.Length--;
            sql.Append(") AND SEQNO in (");
            for (int i = 0; i < str.Count; i++)
            {
                sql.Append(str[i][3] + ",");
            }
            sql.Length--;
            sql.Append(")");
            return sql.ToString();
        }
    }
}
