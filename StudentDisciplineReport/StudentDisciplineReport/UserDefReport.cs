using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StudentDisciplineReport
{
    [ReportTest.MargeClass(Name="報表自訂")]
    public class UserDefReport:ReportTest.framework.MargeGroup
    {
        public System.Data.DataTable BuildMargeData(IEnumerable<string> keys)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("學年度");
            dt.Columns.Add("學期");

            

            return dt;
        }

        public List<string> Fields
        {
            get { return new List<string>(new string []{"學年度","學期"}); }
        }

        public List<string> GroupKeys
        {
            get { return new List<string>();}
        }
    }
}
