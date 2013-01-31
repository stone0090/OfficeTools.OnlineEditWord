using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace OfficeOnline.Demo_DSOframer
{
    public partial class webtest4 : System.Web.UI.Page
    {
        protected string DocUrl
        {
            get { return "http://" + Request.Url.Authority + "/Demo_DSOframer/Doc/111.doc"; }
        }

        protected void Page_Load(object sender, EventArgs e)
        {

        }
    }
}
