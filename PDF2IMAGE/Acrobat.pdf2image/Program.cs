using System;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing.Imaging;
using System.Diagnostics;

namespace Acrobat.pdf2image
{
    public class Program
    {
        /// <summary>
        /// 将PDF文档转换为图片的方法，你可以像这样调用该方法：ConvertPDF2Image("F:\\A.pdf", "F:\\", "A", 0, 0, null, 0);
        /// 因为大多数的参数都有默认值，startPageNum默认值为1，endPageNum默认值为总页数，
        /// imageFormat默认值为ImageFormat.Jpeg，resolution默认值为1
        /// </summary>
        /// <param name="pdfInputPath">PDF文件路径</param>
        /// <param name="imageOutputPath">图片输出路径</param>
        /// <param name="imageName">图片的名字，不需要带扩展名</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换，默认值为1</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换，默认值为PDF总页数</param>
        /// <param name="imageFormat">设置所需图片格式</param>
        /// <param name="resolution">设置图片的分辨率，数字越大越清晰，默认值为1</param>
        public static void ConvertPDF2Image(string pdfInputPath, string imageOutputPath,
            string imageName, int startPageNum, int endPageNum, ImageFormat imageFormat, double resolution)
        {
            Acrobat.CAcroPDDoc pdfDoc = null;
            Acrobat.CAcroPDPage pdfPage = null;
            Acrobat.CAcroRect pdfRect = null;
            Acrobat.CAcroPoint pdfPoint = null;

            // Create the document (Can only create the AcroExch.PDDoc object using late-binding)
            // Note using VisualBasic helper functions, have to add reference to DLL
            pdfDoc = (Acrobat.CAcroPDDoc)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.PDDoc", "");

            // validate parameter
            if (!pdfDoc.Open(pdfInputPath)) { throw new FileNotFoundException(); }
            if (!Directory.Exists(imageOutputPath)) { Directory.CreateDirectory(imageOutputPath); }
            if (startPageNum <= 0) { startPageNum = 1; }
            if (endPageNum > pdfDoc.GetNumPages() || endPageNum <= 0) { endPageNum = pdfDoc.GetNumPages(); }
            if (startPageNum > endPageNum) { int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum; }
            if (imageFormat == null) { imageFormat = ImageFormat.Jpeg; }
            if (resolution <= 0) { resolution = 1; }

            // start to convert each page
            for (int i = startPageNum; i <= endPageNum; i++)
            {
                pdfPage = (Acrobat.CAcroPDPage)pdfDoc.AcquirePage(i - 1);
                pdfPoint = (Acrobat.CAcroPoint)pdfPage.GetSize();
                pdfRect = (Acrobat.CAcroRect)Microsoft.VisualBasic.Interaction.CreateObject("AcroExch.Rect", "");

                int imgWidth = (int)((double)pdfPoint.x * resolution);
                int imgHeight = (int)((double)pdfPoint.y * resolution);

                pdfRect.Left = 0;
                pdfRect.right = (short)imgWidth;
                pdfRect.Top = 0;
                pdfRect.bottom = (short)imgHeight;

                // Render to clipboard, scaled by 100 percent (ie. original size)
                // Even though we want a smaller image, better for us to scale in .NET
                // than Acrobat as it would greek out small text
                pdfPage.CopyToClipboard(pdfRect, 0, 0, (short)(100 * resolution));

                IDataObject clipboardData = Clipboard.GetDataObject();

                if (clipboardData.GetDataPresent(DataFormats.Bitmap))
                {
                    Bitmap pdfBitmap = (Bitmap)clipboardData.GetData(DataFormats.Bitmap);
                    pdfBitmap.Save(Path.Combine(imageOutputPath, imageName) + i.ToString() + "." + imageFormat.ToString(), imageFormat);
                    pdfBitmap.Dispose();
                }
            }

            pdfDoc.Close();
            Marshal.ReleaseComObject(pdfPage);
            Marshal.ReleaseComObject(pdfRect);
            Marshal.ReleaseComObject(pdfDoc);
            Marshal.ReleaseComObject(pdfPoint);

        }

        [STAThread]
        public static void Main(string[] args)
        {
            ConvertPDF2Image("F:\\Events.pdf", "F:\\", "A", 0, 0, null, 0);
        }

    }
}
