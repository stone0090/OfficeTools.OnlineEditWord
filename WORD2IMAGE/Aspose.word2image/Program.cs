using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing.Imaging;
using Aspose.Words;
using System.IO;
using Aspose.Words.Saving;

namespace Aspose.Word2Image
{
    class Program
    {
        /// <summary>
        /// 将Word文档转换为图片的方法（该方法基于第三方DLL），你可以像这样调用该方法：
        /// ConvertPDF2Image("F:\\PdfFile.doc", "F:\\", "ImageFile", 1, 20, ImageFormat.Png, 256);
        /// </summary>
        /// <param name="pdfInputPath">Word文件路径</param>
        /// <param name="imageOutputPath">图片输出路径，如果为空，默认值为Word所在路径</param>
        /// <param name="imageName">图片的名字，不需要带扩展名，如果为空，默认值为Word的名称</param>
        /// <param name="startPageNum">从PDF文档的第几页开始转换，如果为0，默认值为1</param>
        /// <param name="endPageNum">从PDF文档的第几页开始停止转换，如果为0，默认值为Word总页数</param>
        /// <param name="imageFormat">设置所需图片格式，如果为null，默认格式为PNG</param>
        /// <param name="resolution">设置图片的像素，数字越大越清晰，如果为0，默认值为128，建议最大值不要超过1024</param>
        public static void ConvertWordToImage(string wordInputPath, string imageOutputPath,
            string imageName, int startPageNum, int endPageNum, ImageFormat imageFormat, float resolution)
        {
            try
            {
                // open word file
                Aspose.Words.Document doc = new Aspose.Words.Document(wordInputPath);

                // validate parameter
                if (doc == null) { throw new Exception("Word文件无效或者Word文件被加密！"); }
                if (imageOutputPath.Trim().Length == 0) { imageOutputPath = Path.GetDirectoryName(wordInputPath); }
                if (!Directory.Exists(imageOutputPath)) { Directory.CreateDirectory(imageOutputPath); }
                if (imageName.Trim().Length == 0) { imageName = Path.GetFileNameWithoutExtension(wordInputPath); }
                if (startPageNum <= 0) { startPageNum = 1; }
                if (endPageNum > doc.PageCount || endPageNum <= 0) { endPageNum = doc.PageCount; }
                if (startPageNum > endPageNum) { int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = startPageNum; }
                if (imageFormat == null) { imageFormat = ImageFormat.Png; }
                if (resolution <= 0) { resolution = 128; }

                ImageSaveOptions imageSaveOptions = new ImageSaveOptions(GetSaveFormat(imageFormat));
                imageSaveOptions.Resolution = resolution;

                // start to convert each page
                for (int i = startPageNum; i <= endPageNum; i++)
                {
                    imageSaveOptions.PageIndex = i - 1;
                    doc.Save(Path.Combine(imageOutputPath, imageName) + "_" + i.ToString() + "." + imageFormat.ToString(), imageSaveOptions);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static SaveFormat GetSaveFormat(ImageFormat imageFormat)
        {
            SaveFormat sf = SaveFormat.Unknown;
            if (imageFormat.Equals(ImageFormat.Png))
                sf = SaveFormat.Png;
            else if (imageFormat.Equals(ImageFormat.Jpeg))
                sf = SaveFormat.Jpeg;
            else if (imageFormat.Equals(ImageFormat.Tiff))
                sf = SaveFormat.Tiff;
            else if (imageFormat.Equals(ImageFormat.Bmp))
                sf = SaveFormat.Bmp;
            else
                sf = SaveFormat.Unknown;
            return sf;
        }

        static void Main(string[] args)
        {
            ConvertWordToImage("F:\\111.doc","","",0,0,null,0);
        }
    }
}
