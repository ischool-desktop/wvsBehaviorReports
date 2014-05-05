using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace StudentDisciplineReport
{
    public partial class frmReport : FISCA.Presentation.Controls.BaseForm
    {
        public frmReport()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // 讀取選擇學生
            List<string> studentIDList = new List<string>();

            // 傳入敘述
            System.Text.StringBuilder sb = new StringBuilder();
            sb.AppendLine("學校資訊");
            sb.AppendLine("學生基本資料");
            sb.AppendLine("學生缺曠明細;學年度=" + iptSchoolYear.Value + ",學期=" + iptSemester.Value);
            ReportTest.Config conf = new ReportTest.Config(sb.ToString());
            ReportTest.Marge m = new ReportTest.Marge(conf, studentIDList);
            m.MargeData();
            System.Data.DataTable dt = m.GetDataTable();

            // 讀取樣版
            Aspose.Words.Document doc = new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.學生個人獎懲紀錄表_世界高中));
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);

            // 合併
            doc.MailMerge.DeleteFields();

            string filePath = System.Windows.Forms.Application.StartupPath + "\\test獎懲報表所有學年度學期.doc";
            doc.Save(filePath, Aspose.Words.SaveFormat.Doc);
            System.Diagnostics.Process.Start(filePath);

        }

        private void frmReport_Load(object sender, EventArgs e)
        {
            // 預設學年度學期
            iptSchoolYear.Value=int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value=int.Parse(K12.Data.School.DefaultSemester);
        }

    }
}
