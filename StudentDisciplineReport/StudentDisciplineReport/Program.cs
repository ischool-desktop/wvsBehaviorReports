using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StudentDisciplineReport
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void Main()
        {
            FISCA.Permission.Catalog cat = FISCA.Permission.RoleAclSource.Instance["學生"]["功能按鈕"];
            cat.Add(new FISCA.Permission.RibbonFeature("ischool.StudentDisciplineReport.form", "世界高中學生個人獎懲紀錄表"));
            cat.Add(new FISCA.Permission.RibbonFeature("ischool.StudentDisciplineReport.form1", "世界高中學生個人缺曠獎懲明細表"));


            var btn =K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["世界高中學生個人獎懲紀錄表"];
            btn.Enable = FISCA.Permission.UserAcl.Current["ischool.StudentDisciplineReport.form"].Executable;
            btn.Click += delegate {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    frmReport fr = new frmReport();
                    fr.ShowDialog();
                }
            };

            var btn1 = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["學務相關報表"]["世界高中學生個人缺曠獎懲明細表"];
            btn1.Enable = FISCA.Permission.UserAcl.Current["ischool.StudentDisciplineReport.form1"].Executable;
            btn1.Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    frmReport1 fr = new frmReport1();
                    fr.ShowDialog();
                }
            };



            // // 測試按鈕
            //K12.Presentation.NLDPanels.Student.ListPaneContexMenu["世界高中學生個人獎懲紀錄表"].Enable = true;
            //K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test世界高中學生個人獎懲紀錄表"].Click += delegate
            //{
            //    if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
            //    {
            //        frmReport fr = new frmReport();
            //        fr.ShowDialog();
            //    }
            //};


            //// 測試按鈕
            //K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test世界高中學生個人缺曠獎懲明細表"].Enable = true;
            //K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test世界高中學生個人缺曠獎懲明細表"].Click += delegate
            //{
            //    if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
            //    {
            //        frmReport1 fr = new frmReport1();
            //        fr.ShowDialog();
            //    }
            //};
        }
    }
}
