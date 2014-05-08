using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace StudentDisciplineReport
{
    public partial class frmReport1 : FISCA.Presentation.Controls.BaseForm
    {
        List<Config> _ConfigList;
        FISCA.UDT.AccessHelper _AccessHelper;
        string _Name = "世界高中學生個人缺曠獎懲明細表";

        public frmReport1()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            btnPrint.Enabled = false;

            try
            {
                // 讀取選擇學生
                List<string> studentIDList = K12.Presentation.NLDPanels.Student.SelectedSource;

                // 傳入敘述
                System.Text.StringBuilder sb = new StringBuilder();
                sb.AppendLine("學校資訊");
                sb.AppendLine("學生基本資料");
                sb.AppendLine("報表自訂;學年度=" + iptSchoolYear.Value + ",學期=" + iptSemester.Value);
                sb.AppendLine("獎懲明細;學年度=" + iptSchoolYear.Value + ",學期=" + iptSemester.Value);
                sb.AppendLine("缺曠明細;學年度=" + iptSchoolYear.Value + ",學期=" + iptSemester.Value);
                sb.AppendLine("獎懲統計;學年度=" + iptSchoolYear.Value);
                sb.AppendLine("缺曠統計;學年度=" + iptSchoolYear.Value);
                sb.AppendLine("獎懲統計_功過相抵;學年度=" + iptSchoolYear.Value);
                ReportTest.Config conf = new ReportTest.Config(sb.ToString());

                ReportTest.Marge m = new ReportTest.Marge(conf, studentIDList);
                m.AddType(typeof(UserDefReport));

                m.MargeData();
                System.Data.DataTable dt = m.GetDataTable();

                // 讀取樣版
                Aspose.Words.Document doc = _ConfigList[0].GetTemplate();//new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.學生個人缺曠獎懲明細表_世界高中));
                Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
                doc.MailMerge.Execute(dt);
                // 合併
                doc.MailMerge.DeleteFields();

                string filePath = System.Windows.Forms.Application.StartupPath + "\\世界高中學生個人缺曠獎懲明細表.doc";
                // 輸出至Word 
                Utility.ExportDoc(filePath, doc);

                btnPrint.Enabled = true;
            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤," + ex.Message);
                btnPrint.Enabled = true;
            }
        }

        private void frmReport_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            // 預設學年度學期
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);
            _AccessHelper = new FISCA.UDT.AccessHelper();
            _ConfigList = new List<Config>();

            LoadTemplate();
        }

        private void LoadTemplate()
        {
            _ConfigList.Clear();
            _ConfigList = _AccessHelper.Select<Config>("Name='" + _Name + "'");
            if (_ConfigList.Count == 0)
            {
                Config conf = new Config();
                conf.Name = _Name;
                conf.SaveTemplate(conf.GetTemplate());
                _ConfigList.Add(conf);
                _ConfigList.SaveAll();
            }
        }

        private void lnkDefault_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 檢視套印樣版                      

            string reportName = _ConfigList[0].Name;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {

                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                //this.Configure.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
                _ConfigList[0].GetTemplate().Save(stream, Aspose.Words.SaveFormat.Doc);
                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        //document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.學生個人獎懲紀錄表_世界高中, 0, Properties.Resources.學生個人獎懲紀錄表_世界高中.Length);
                        stream.Flush();
                        stream.Close();

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
        }

        private void lnkUserDef_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // 變更套印樣版            
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Excel檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    Aspose.Words.Document doc = new Aspose.Words.Document(dialog.FileName);
                    _ConfigList[0].SaveTemplate(doc);
                    _ConfigList.SaveAll();
                    FISCA.Presentation.Controls.MsgBox.Show("上傳樣版完成");
                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }      
        }
    }
}
