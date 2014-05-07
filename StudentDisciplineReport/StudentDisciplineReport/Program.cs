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
             // 測試按鈕
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test世界高中學生個人獎懲紀錄表"].Enable = true;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test世界高中學生個人獎懲紀錄表"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    frmReport fr = new frmReport();
                    fr.ShowDialog();
                }
            };


            // 測試按鈕
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test世界高中學生個人缺曠獎懲明細表"].Enable = true;
            K12.Presentation.NLDPanels.Student.ListPaneContexMenu["Test世界高中學生個人缺曠獎懲明細表"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0)
                {
                    frmReport1 fr = new frmReport1();
                    fr.ShowDialog();
                }
            };
        }
    }
}
