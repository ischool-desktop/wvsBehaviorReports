using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StudentDisciplineReport
{
    [ReportTest.MargeClass(Name="報表自訂")]
    public class UserDefReport:ReportTest.framework.MargeGroup
    {
        [ReportTest.MargeField(FieldName="學年度",FieldType="Int")]
        public int? SchoolYear;

        [ReportTest.MargeField(FieldName="學期",FieldType="Int")]
        public int? Semester;

        public System.Data.DataTable BuildMargeData(IEnumerable<string> keys)
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add("學年度");
            dt.Columns.Add("學期");

            if(SchoolYear.HasValue && Semester.HasValue)
            foreach (string sid in keys)
            {
                dt.Rows.Add(
                    sid
                    ,SchoolYear.Value
                    ,Semester.Value
                    );
            
            }            

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
