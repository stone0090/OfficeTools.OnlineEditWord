using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.IO;
using System.Drawing.Imaging;

namespace TallComponents.PDF.Rasterizer.pdf2image
{
    public static class Program
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
            FileStream fs = new FileStream(pdfInputPath, FileMode.Open);

            TallComponents.PDF.Rasterizer.Document doc = new TallComponents.PDF.Rasterizer.Document(fs);

            if (!Directory.Exists(imageOutputPath))
            {
                Directory.CreateDirectory(imageOutputPath);
            }

            // validate pageNum
            if (startPageNum <= 0)
            {
                startPageNum = 1;
            }

            if (endPageNum > doc.Pages.Count)
            {
                endPageNum = doc.Pages.Count;
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
                using (FileStream fs1 = File.Create(imageOutputPath + @"~temp" + i + ".tmp"))
                {
                    TallComponents.PDF.Rasterizer.ConvertToTiffOptions option = new TallComponents.PDF.Rasterizer.ConvertToTiffOptions();
                    doc.Pages[i].ConvertToTiff(fs1, option);
                    System.Drawing.Image img = System.Drawing.Image.FromStream(fs1);
                    img.Save(imageOutputPath + imageName + i.ToString() + "." + imageFormat.ToString(), imageFormat);
                }

                File.Delete(imageOutputPath + @"~temp" + i + ".tmp");
            }

            fs.Dispose();
        }

        public static void Main(string[] args)
        {
            ConvertPDF2Image("F:\\Events.pdf", "F:\\", "A", 1, 5, ImageFormat.Jpeg, Definition.One);
        }

    }
}
