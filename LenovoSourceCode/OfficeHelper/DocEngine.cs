using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Words;
using Aspose.Words.Rendering;
using System.Drawing;
using System.Drawing.Imaging;

namespace LenovoCW.MOA
{
    public class DocEngine
    {
        /// <summary>
        /// 裁剪的边框宽度
        /// </summary>
        const int CutBorderWidth = 1;//60;
        /// <summary>
        /// 裁剪的顶部高度
        /// </summary>
        const int CutTopHeight = 1;//60;
        /// <summary>
        /// 裁剪的底部高度
        /// </summary>
        const int CutBottomHeight = 1;//60;

        private int _imageHeight;
        private int _imageWidth;

        /// <summary>
        /// 转换DOC文件至图片
        /// </summary>
        /// <param name="srcPath"></param>
        /// <param name="destDir"></param>
        /// <param name="fileGUID"></param>
        /// <param name="pageIndex"></param>
        /// <param name="imgQuality"></param>
        /// <returns></returns>
        public string Convert(string srcPath, string destDir, string fileGUID, Boolean combineImages, int pageIndex = -1, int imgQuality = 0)
        {
            string imgPaths = destDir + fileGUID;

            if (!Directory.Exists(imgPaths))
            {
                Directory.CreateDirectory(imgPaths);
            }


            Document doc = new Document(srcPath);
            try
            {
                int pageCount = doc.PageCount;
                return ProcessPage(doc, destDir, fileGUID, pageCount, imgQuality,combineImages);
            }
            catch (Exception)
            {
                return "";
            }
        }

        private string ProcessPage(Document doc, string destDir, string fileGUID, int pageCount, int imgQuality,Boolean combineImages)
        {
            string imgPaths = destDir + fileGUID;//Guid.NewGuid().ToString();
            if (!Directory.Exists(imgPaths))
            {
                Directory.CreateDirectory(imgPaths);
            }

            if (imgQuality == 0)
            {
                imgQuality = 100;
            }

            var result = "";

            ImageOptions imageOptions = new ImageOptions();
            imageOptions.JpegQuality = 100;

            var tifImagePath = string.Format(@"{0}\{1}.tif", imgPaths, fileGUID);
            doc.SaveToImage(0, doc.PageCount, tifImagePath, imageOptions);
            Utils.Tif2Jpeg(tifImagePath, imgQuality, false);

            //拼接图片
            if (combineImages)
            {
                int imgCount = pageCount;
                if (imgCount > 0)
                {
                    Bitmap resultImg = null;
                    Graphics resultGraphics = null;
                    Image tempImage;

                    for (int i = 0; i < imgCount; i++)
                    {
                        var imgFile = imgPaths + "\\" + i + ".jpg";
                        if (File.Exists(imgFile))
                        {
                            tempImage = Image.FromFile(imgFile);

                            if (resultImg == null)
                            {
                                _imageHeight = tempImage.Height;

                                resultImg = new Bitmap(tempImage.Width, tempImage.Height * imgCount);
                                resultGraphics = Graphics.FromImage(resultImg);


                            }


                            if (i == 0)
                            {
                                resultGraphics.DrawImage(tempImage, 0, 0);
                            }
                            else
                            {
                                resultGraphics.DrawImage(tempImage, 0, _imageHeight * i);
                            }

                            tempImage.Dispose();
                        }
                    }
                    result = string.Format(@"{0}\{1}.jpg", imgPaths, fileGUID);
                    ImageUtility.CompressAsJPG(resultImg, result, imgQuality);
                    resultGraphics.Dispose();
                }


            }
            else
            {
                result = string.Format(@"{0}\{1}.jpg", imgPaths, 0);
            }
            return result;
               
        }
    }

}
