using System.Drawing.Imaging;
using System.IO;
using PDFLibNet;

namespace PDFLibNet.pdf2image
{
    class Program
    {

        public enum Definition
        {
            One = 1, Two = 2, Three = 3, Four = 4, Five = 5, Six = 6, Seven = 7, Eight = 8, Nine = 9, Ten = 10
        }

        /// <summary>
        /// 将PDF文档转换为图片的方法
        /// </summary>
        /// <param name="pdfInputPath">PDF文件路径</param>
        /// <param name="imageOutputPath">图片输出路径</param>
        /// <param name="imageName">生成图片的名字</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换</param>
        /// <param name="imageFormat">设置所需图片格式</param>
        /// <param name="definition">设置图片的清晰度，数字越大越清晰</param>
        public static void ConvertPDF2Image(string pdfInputPath, string imageOutputPath,
            string imageName, int startPageNum, int endPageNum, ImageFormat imageFormat, Definition definition)
        {
            PDFWrapper pdfWrapper = new PDFWrapper();

            pdfWrapper.LoadPDF(pdfInputPath);

            if (!System.IO.Directory.Exists(imageOutputPath))
            {
                Directory.CreateDirectory(imageOutputPath);
            }

            // validate pageNum
            if (startPageNum <= 0)
            {
                startPageNum = 1;
            }

            if (endPageNum > pdfWrapper.PageCount)
            {
                endPageNum = pdfWrapper.PageCount;
            }

            if (startPageNum > endPageNum)
            {
                int tempPageNum = startPageNum;
                startPageNum = endPageNum;
                endPageNum = startPageNum;
            }

            // start to convert each page
            for (int i = startPageNum; i <= endPageNum; i++)
            {
                pdfWrapper.ExportJpg(imageOutputPath + imageName + i.ToString() + ".jpg", i, i, 150, 90);//这里可以设置输出图片的页数、大小和图片质量
                if (pdfWrapper.IsJpgBusy) { System.Threading.Thread.Sleep(50); }
            }

            pdfWrapper.Dispose();
        }

        public static void Main(string[] args)
        {
            ConvertPDF2Image("F:\\Events.pdf", "F:\\", "A", 1, 5, ImageFormat.Jpeg, Definition.One);
        }

    }
}
