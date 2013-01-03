using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WebSupergoo.ABCpdf6;

namespace ABCpdf.pdf2image
{
    class Program
    {
        static void Main(string[] args)
        {
            Doc theDoc = new Doc();
            //theDoc.Read(Server.MapPath("../Rez/spaceshuttle.pdf"));
            theDoc.Read("F:\\Events.pdf");

            theDoc.Rendering.DotsPerInch = 200;
            
            for (int i = 0; i < 4; i++)
            {
                theDoc.PageNumber = i;
                theDoc.Rect.String = theDoc.CropBox.String;
                theDoc.Rendering.Save("F:\\Events" + i.ToString() + ".JPG");
            }
        }
    }
}
