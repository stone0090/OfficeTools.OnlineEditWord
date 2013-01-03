using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using O2S.Components.PDFRender4NET;

namespace LenovoCW.MOA
{
    public class PDFEngine
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

        public bool ThumbnailCallback()
        {
            return false;
        }

        /// <summary>
        /// 转换PDF文件
        /// </summary>
        /// <param name="srcPath"></param>
        /// <param name="destDir"></param>
        /// <param name="fileGUID"></param>
        /// <param name="pageIndex">转换的页码,-1表示转换所有页</param>
        /// <param name="imgQuality"></param>
        /// <param name="combineImages"></param>
        /// <returns></returns>
        public string Convert(string srcPath, string destDir,string fileGUID,Boolean combineImages,int pageIndex=-1,int imgQuality=0)
        {
            string imgPaths = destDir + fileGUID;//Guid.NewGuid().ToString();
            string result = string.Empty;
            if (pageIndex < 0)
            {
                if (combineImages)
                    result = string.Format(@"{0}\{1}.jpg", imgPaths, fileGUID);
                else
                    result = string.Format(@"{0}\0.jpg", imgPaths);
            }
            else
            {
                result = string.Format(@"{0}\{1}.jpg", imgPaths, pageIndex);
            }
            //如果已经存在，则直接返回
            if (File.Exists(result))
                return result;


            if (!Directory.Exists(imgPaths))
            {
                Directory.CreateDirectory(imgPaths);
            }

            
            PDFFile pdf = PDFFile.Open(srcPath);

            try
            {
                int pageCount = pdf.PageCount;
                //根据页索取
                if (pageIndex >= 0 && pageIndex < pageCount)
                {
                    return ProcessPage(pdf, destDir, fileGUID, pageIndex, pageCount, imgQuality);
                }
                //获取所有页
                else if (pageIndex == -1)
                {

                    string[] ImgPaths = new string[pdf.PageCount];

                    //全部取的情况
                    for (int i = 0; i < pdf.PageCount; i++)
                    {
                        ImgPaths[i] = ProcessPage(pdf, destDir, fileGUID, i, pageCount, imgQuality);

                    }

                    //拼接图片
                    if (combineImages)
                    {
                        int imgCount = ImgPaths.Count();
                        if (imgCount > 0)
                        {
                            Bitmap resultImg = new Bitmap(_imageWidth, _imageHeight * imgCount);
                            Graphics resultGraphics = Graphics.FromImage(resultImg);
                            Image tempImage;
                            for (int i = 0; i < ImgPaths.Length; i++)
                            {
                                tempImage = Image.FromFile(ImgPaths[i]);
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

                            ImageUtility.CompressAsJPG(resultImg, result, imgQuality);
                            resultGraphics.Dispose();

                        }
                    }

                    return result;
                }
                return "";
            }
            finally
            {
                pdf.Dispose();
            }
        }

        private string ProcessPage(PDFFile pdf, string destDir, string guid, int pageIndex, int pageCount, int imgQuality)
        {
            string imgPaths = destDir + guid;//Guid.NewGuid().ToString();
            if (!Directory.Exists(imgPaths))
            {
                Directory.CreateDirectory(imgPaths);
            }

            if (imgQuality == 0)
            {
                imgQuality = 100;
            }

            Bitmap oriBmp = pdf.GetPageImage(pageIndex, 96);

            Bitmap bmp = ImageUtility.CutAsBmp(oriBmp, CutBorderWidth, CutTopHeight, oriBmp.Width - 2 * CutBorderWidth, oriBmp.Height - CutTopHeight - CutBottomHeight);

            string result = string.Format(@"{0}\{1}.jpg", imgPaths, pageIndex);

            if (bmp.Width >= 700)
            {
                _imageHeight = (int)bmp.Height;// / 2;
                _imageWidth = (int)bmp.Width ;// / 2;
                ImageUtility.ThumbAsJPG(bmp, result, _imageWidth, _imageHeight, imgQuality);

                //tempImg = bmp.GetThumbnailImage((int)bmp.Width / 2, (int)bmp.Height / 2, new Image.GetThumbnailImageAbort(ThumbnailCallback), IntPtr.Zero);
                //tempImg.Save(string.Format(@"{0}\{1}-ori.jpg", imgPaths, i), System.Drawing.Imaging.ImageFormat.Jpeg);
            }
            else
            {

                _imageHeight = bmp.Height;
                _imageWidth = bmp.Width;

                ImageUtility.CompressAsJPG(bmp, result, imgQuality);


            }
            return result;
            //bmp.Save(string.Format(@"{0}\{1}.jpg",imgPaths,i), System.Drawing.Imaging.ImageFormat.Jpeg);
        }
    }

}
