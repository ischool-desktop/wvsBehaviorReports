using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StudentDisciplineReport
{
    [FISCA.UDT.TableName("ischool.StudentDisciplineReport.Config")]
    public class Config:FISCA.UDT.ActiveRecord
    {
        [FISCA.UDT.Field]
        public string Name { get; set; }

        [FISCA.UDT.Field]
        public string TemplateStream { get; set; }


        public Aspose.Words.Document GetTemplate()
        {
            Aspose.Words.Document doc = null;

            if (!string.IsNullOrEmpty(this.TemplateStream))
            {
                // 取的樣版
                doc = new Aspose.Words.Document(new System.IO.MemoryStream(Convert.FromBase64String(this.TemplateStream)));
            }
            if (doc == null)
            {
                if (Name == "世界高中學生個人獎懲紀錄表")
                    doc = new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.學生個人獎懲紀錄表_世界高中));
                else
                    doc = new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.學生個人缺曠獎懲明細表_世界高中));
            }
            return doc;
        }

        public void SaveTemplate(Aspose.Words.Document doc)
        {
            System.IO.MemoryStream stream = new System.IO.MemoryStream();
            doc.Save(stream, Aspose.Words.SaveFormat.Doc);
            this.TemplateStream = Convert.ToBase64String(stream.ToArray());
            this.Save();
        }
    }
}
