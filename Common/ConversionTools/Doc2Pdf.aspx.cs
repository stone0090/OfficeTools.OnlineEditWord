using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Aspose.Words;

namespace WebOffice
{
    /// <summary>
    /// Aspose.Words的Doc转换Pdf的功能很差，字体丢失，格式错乱（不建议使用）
    /// </summary>
    public partial class Doc2Pdf : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            if (this.FileUpload1.HasFile)
            {
                using (Stream stream = this.FileUpload1.PostedFile.InputStream)
                {
                    var doc = new Aspose.Words.Document(stream, this.FileUpload1.PostedFile.FileName);

                    var filePath = Page.MapPath("temp");
                    if (!Directory.Exists(filePath))
                        Directory.CreateDirectory(filePath);

                    var fileName = Path.Combine(filePath, Guid.NewGuid().ToString() + ".pdf");

                    doc.SaveToPdf(fileName);
                }
            }
        }
    }
}
