using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Data;

namespace DAL
{
    public class Tools
    {
        public static string getSokocd()
        {
            var xml = XDocument.Load(@"..\Mitsui.xml");
            var links = xml.Descendants("sokocd")
                        .Attributes("val")
                        .Select(element => element.Value).ToList();
            var SOKOCD = links[0].ToString();
            return SOKOCD;
        }

        public static string GetKenSu()
        {
            var xml = XDocument.Load(@"..\Mitsui.xml");
            var links = xml.Descendants("search")
                        .Attributes("kensu")
                        .Select(element => element.Value).ToList();
            var kensu = links[0].ToString();
            return kensu;
        }
        
        public static DataTable GetAreanm()
        {
            var sql = "select areanm from M_AREA ma left join M_MNRCODE mm on ma.areacd = mm.code group by ma.areanm,nvl(mm.kbn1,'Z') order by nvl(mm.kbn1,'Z')";
            DEV10G2U dEV = new DEV10G2U();
            var dt = dEV.executeSelectQuery(sql);
            return dt;
        }
    }
}
