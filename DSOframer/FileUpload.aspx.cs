using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Configuration;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DSOframer
{
    public partial class FileUpload : System.Web.UI.Page
    {
        protected string DocUrl
        {
            get { return "http://" + Request.Url.Authority + "/Doc/test.doc"; }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            //if (!IsPostBack) return;
            if (Request.InputStream.Length <= 0) return;
            using (var stream = Request.InputStream)
            {
                using (var fs = new FileStream(Server.MapPath("doc/test.doc"), FileMode.Create))
                {
                    int readCount;
                    var buffer = new byte[1024];
                    while ((readCount = stream.Read(buffer, 0, 1024)) > 0)
                        fs.Write(buffer, 0, readCount);
                    fs.Flush();
                }
            }
        }
    }
}