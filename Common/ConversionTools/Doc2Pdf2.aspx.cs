using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using Microsoft.Office.Interop.Word;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace WebOffice
{
    /// <summary>
    /// office的Doc转换Pdf的功能不错，缺点服务器需要安装office07软件和Pdf插件，itextSharp加文字水印和图片水印功能
    /// </summary>
    public partial class Doc2Pdf2 : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            if (this.FileUpload1.HasFile)
            {
                var filePath = Page.MapPath("temp");
                if (!Directory.Exists(filePath))
                    Directory.CreateDirectory(filePath);

                var oPdfName = Guid.NewGuid().ToString() + ".pdf";
                var tPdfName = Guid.NewGuid().ToString() + ".pdf";
                var oPdfFullName = Path.Combine(filePath, oPdfName);
                var tPdfFullName = Path.Combine(filePath, tPdfName);

                ConvertWord2PDF(this.FileUpload1.PostedFile.FileName, filePath, oPdfName);

                var dllPath = Page.MapPath("dll");
                var iTextAsian = Path.Combine(dllPath, "iTextAsian.dll");
                var iTextAsianCmaps = Path.Combine(dllPath, "iTextAsianCmaps.dll");

                PDFTextWatermark(iTextAsian, iTextAsianCmaps, oPdfFullName, tPdfFullName, "中国南方航空股份有限公司");

                Response.TransmitFile(Path.Combine(filePath, tPdfFullName));
            }
        }

        public static void ConvertWord2PDF(string wordInputPath, string pdfOutputPath, string pdfName)
        {
            try
            {
                string pdfExtension = ".pdf";

                // validate patameter
                if (!Directory.Exists(pdfOutputPath)) { Directory.CreateDirectory(pdfOutputPath); }
                if (pdfName.Trim().Length == 0) { pdfName = Path.GetFileNameWithoutExtension(wordInputPath); }
                if (!(Path.GetExtension(pdfName).ToUpper() == pdfExtension.ToUpper())) { pdfName = pdfName + pdfExtension; }

                object paramSourceDocPath = wordInputPath;
                object paramMissing = Type.Missing;

                string paramExportFilePath = Path.Combine(pdfOutputPath, pdfName);

                // create a word application object
                ApplicationClass wordApplication = new ApplicationClass();
                Microsoft.Office.Interop.Word.Document wordDocument = null;

                //set exportformat to pdf 
                WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
                bool paramOpenAfterExport = false;
                WdExportOptimizeFor paramExportOptimizeFor = WdExportOptimizeFor.wdExportOptimizeForPrint;
                WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
                int paramStartPage = 0;
                int paramEndPage = 0;
                WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
                bool paramIncludeDocProps = true;
                bool paramKeepIRM = true;
                WdExportCreateBookmarks paramCreateBookmarks = WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                bool paramDocStructureTags = true;
                bool paramBitmapMissingFonts = true;
                bool paramUseISO19005_1 = false;

                try
                {
                    // Open the source document.
                    wordDocument = wordApplication.Documents.Open(
                        ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing, ref paramMissing, ref paramMissing,
                        ref paramMissing);

                    // Export it in the specified format.
                    if (wordDocument != null)
                        wordDocument.ExportAsFixedFormat(paramExportFilePath,
                            paramExportFormat, paramOpenAfterExport,
                            paramExportOptimizeFor, paramExportRange, paramStartPage,
                            paramEndPage, paramExportItem, paramIncludeDocProps,
                            paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                            paramBitmapMissingFonts, paramUseISO19005_1,
                            ref paramMissing);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    // Close and release the Document object.
                    if (wordDocument != null)
                    {
                        wordDocument.Close(ref paramMissing, ref paramMissing, ref paramMissing);
                        wordDocument = null;
                    }

                    // Quit Word and release the ApplicationClass object.
                    if (wordApplication != null)
                    {
                        wordApplication.Quit(ref paramMissing, ref paramMissing, ref paramMissing);
                        wordApplication = null;
                    }

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /// <summary>
        /// 创建PDF的文字水印
        /// </summary>
        /// <param name="iTextAsian_Path">iTextAsian.dll所在路径</param>
        /// <param name="iTextAsianCmaps_Path">iTextAsianCmaps.dll所在路径</param>
        /// <param name="inputfilepath">源PDF文件</param>
        /// <param name="outputfilepath">加水印后PDF文件</param>
        /// <param name="waterMarkText">文字水印文本</param>
        public static void PDFTextWatermark(string iTextAsian_Path, string iTextAsianCmaps_Path, string inputfilepath, string outputfilepath, string waterMarkText)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            pdfReader = new PdfReader(inputfilepath);
            //pdf总的页数
            int numberOfPages = pdfReader.NumberOfPages;
            iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
            float width = psize.Width;
            float height = psize.Height;
            pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));

            PdfContentByte content;

            //设置字体，载入亚洲字体资源，无此操作的话，不能显示包括中文、日文、韩文等内容   
            BaseFont.AddToResourceSearch(iTextAsianCmaps_Path);
            BaseFont.AddToResourceSearch(iTextAsian_Path);
            BaseFont bf = BaseFont.CreateFont("STSong-Light", "UniGB-UCS2-H", BaseFont.NOT_EMBEDDED);


            //水印文字，文字长度   
            int j = waterMarkText.Length;

            //循环对每页插入水印   
            for (int i = 1; i <= numberOfPages; i++)
            {

                content = pdfStamper.GetUnderContent(i);
                //开始   
                content.BeginText();
                //设置颜色   
                content.SetColorFill(iTextSharp.text.BaseColor.GRAY);
                //设置字体及字号   
                content.SetFontAndSize(bf, 18);
                //设置起始位置   
                content.SetTextMatrix(20, 25);
                //设置对齐以及倾斜方式等
                content.ShowTextAligned((int)(Element.ALIGN_CENTER), waterMarkText, content.PdfDocument.PageSize.Width / 2, content.PdfDocument.PageSize.Height / 2 + 100, 60f);
                //关闭
                content.EndText();

            }
            pdfStamper.Close();

        }

        /// <summary> 
        ///  PDF加水印 
        /// </summary> 
        /// <param name="inputfilepath">源PDF文件</param> 
        /// <param name="outputfilepath">加水印后PDF文件 </param> 
        /// <param name="ModelPicName">水印文件路径</param> 
        /// <param name="top">离顶部距离</param> 
        /// <param name="left">离左边距离,如果为负,则为离右边距离</param> 
        /// <param name="strMsg">返回信息</param> 
        /// <returns>返回</returns> 
        public static bool PDFWatermark(string inputfilepath, string outputfilepath, string ModelPicName, float top, float left, ref string strMsg)
        {
            PdfReader pdfReader = null;
            PdfStamper pdfStamper = null;
            try
            {
                pdfReader = new PdfReader(inputfilepath);
                int numberOfPages = pdfReader.NumberOfPages;
                iTextSharp.text.Rectangle psize = pdfReader.GetPageSize(1);
                float width = psize.Width;
                float height = psize.Height;
                pdfStamper = new PdfStamper(pdfReader, new FileStream(outputfilepath, FileMode.Create));
                PdfContentByte waterMarkContent;
                iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(ModelPicName);
                image.GrayFill = 20;//透明度，灰色填充
                //image.Rotation//旋转
                //image.RotationDegrees//旋转角度
                //水印的位置 
                if (left < 0)
                {
                    left = width - image.Width + left;
                }
                image.SetAbsolutePosition(left, (height - image.Height) - top);

                //每一页加水印,也可以设置某一页加水印 
                for (int i = 1; i <= numberOfPages; i++)
                {
                    waterMarkContent = pdfStamper.GetUnderContent(i);
                    waterMarkContent.AddImage(image);
                }
                strMsg = "success";
                return true;
            }
            catch (Exception ex)
            {
                strMsg = ex.Message.Trim();
                return false;
            }
            finally
            {
                if (pdfStamper != null)
                    pdfStamper.Close();
                if (pdfReader != null)
                    pdfReader.Close();
            }
        }   
    
    }
}
