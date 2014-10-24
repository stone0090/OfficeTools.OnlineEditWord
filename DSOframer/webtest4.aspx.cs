using System;

namespace DSOframer
{
    public partial class webtest4 : System.Web.UI.Page
    {
        protected string DocUrl
        {
            get { return "http://" + Request.Url.Authority + "/Doc/test.doc"; }
        }

        protected void Page_Load(object sender, EventArgs e)
        {

        }
    }
}
