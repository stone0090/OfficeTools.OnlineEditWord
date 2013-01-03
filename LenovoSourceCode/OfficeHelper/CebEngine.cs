using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;
using System.Threading;

namespace LenovoCW.MOA
{
    /// <summary>
    /// CEB转换引擎类
    /// </summary>
    public class CebEngine
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
        /// 转换CEB文件
        /// </summary>
        /// <param name="srcFilePath">源文件路径</param>
        /// <param name="destDir">目标目录</param>
        /// <param name="fileGUID">文件标识</param>
        /// <param name="cebImagePath">CEB图片路径</param>
        /// <param name="imgQuality">转换的图片质量</param>
        /// <returns>转换完成后对应的图片路径</returns>
        public string Convert(string srcFilePath, string destDir, string fileGUID, string cebImagePath,Boolean combineImages,int imgQuality = 0)
        {
            string imgSaveDir = destDir + fileGUID;//最终生成的文件目录

            string result = string.Empty;
            if (combineImages)
                result = string.Format(@"{0}\{1}.jpg", imgSaveDir, fileGUID);
            else
                result = string.Format(@"{0}\0.jpg", imgSaveDir);


            if (!Directory.Exists(imgSaveDir))
            {
                Directory.CreateDirectory(imgSaveDir);
            }
            else
            {
                return result;
            }

           
            try
            {
                string tifFilePath = string.Format(@"{0}\{1}\{2}.tif",cebImagePath,fileGUID,fileGUID); 
                //if (!File.Exists(tifFilePath))
                //{
                //    //执行转换
                //ShellExecute(IntPtr.Zero, "", cebAppPath, @"/p " + srcFilePath, "", 0);
                //    System.Diagnostics.Process p = new System.Diagnostics.Process();
                //    p.StartInfo.FileName = cebAppPath;//需要启动的程序名       
                //    p.StartInfo.Arguments = @"/p " + srcFilePath;//启动参数  
                //    p.StartInfo.Verb = "runas";
                //    p.Start();//启动       

                //}

                return ProcessPage(tifFilePath, imgQuality, combineImages);
                
            }
            catch (Exception)
            {
                return "";
            }
        }



        private string ProcessPage(string tifFilePath, int imgQuality, Boolean combineImages)
        {

            if (imgQuality == 0)
            {
                imgQuality = 100;
            }

            var loop = 0;
            //60秒超时
            while (loop<120)
            {
                if (File.Exists(tifFilePath))
                {
                    //确保文件已转换完成
                    try
                    {
                        File.Move(tifFilePath, tifFilePath + ".tmp");
                        File.Move(tifFilePath + ".tmp",tifFilePath);
                        break;
                    }
                    catch (Exception)
                    {

                    }
                }
                Thread.Sleep(500);
                loop++;
            }

            //拼接图片
            return Utils.Tif2Jpeg(tifFilePath, imgQuality, combineImages);               
        }
    }

}
