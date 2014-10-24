using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DSOframer
{
    public partial class FileDownload : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            //Response.AddHeader("Content-Type", "application/msword");
            //Response.AppendHeader("Content-Disposition", String.Format("attachment; filename=\"{0}\"", "test.doc"));
            Response.AddHeader("Content-Type", "octet-stream");
            Response.TransmitFile(Server.MapPath("doc/test.doc"));
            Response.End();
        }
    }
}