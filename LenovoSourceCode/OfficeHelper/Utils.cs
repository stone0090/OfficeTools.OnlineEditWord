using System;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Net;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Checksums;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Drawing;
using System.Drawing.Imaging;
using System.Net.NetworkInformation;

namespace LenovoCW.MOA
{
    public class Utils
    {

        private int _zoomPercent;
        private string _tagWidth;
        private Boolean _firstTag=true;
 
        /// <summary>
        /// 压缩文件
        /// </summary>
        /// <param name="fileName">源文件名称</param>
        /// <param name="ZipFileName">压缩文件名称</param>
        public static void ZipFile(string fileName, string zipFileName)
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                using (FileStream ZipFile = File.Create(zipFileName))
                {
                    ZipOutputStream ZipStream = new ZipOutputStream(ZipFile);
                    ZipEntry ZipEntry = new ZipEntry("ZippedFile");
                    
                    ZipStream.PutNextEntry(ZipEntry);
                    //ZipStream.SetLevel(CompressionLevel);
                    byte[] buffer = new byte[2048];
                    System.Int32 size = fs.Read(buffer, 0, buffer.Length);
                    ZipStream.Write(buffer, 0, size);
                    while (true)
                    {
                        int sizeRead = fs.Read(buffer, 0, buffer.Length);
                        if (sizeRead > 0)
                            ZipStream.Write(buffer, 0, sizeRead);
                        else
                            break;
                    }
                    ZipStream.Flush();
                    ZipStream.Finish();
                }
            }

        }

        /// <summary>
        /// 解压文件
        /// </summary>
        /// <param name="fileName">压缩文件名称</param>
        /// <param name="unzipFileName">解压文件名称</param>
        public static void UnZipFile(string fileName, string unzipFileName)
        {
            try
            {
                //FileStream fss = new FileStream(fileName, FileMode.Open);
                using (FileStream fs = File.OpenRead(fileName))
                {
                    using (ZipInputStream zis = new ZipInputStream(fs))
                    {
                        ZipEntry entry;
                        while ((entry = zis.GetNextEntry()) != null)
                        {
                            using (FileStream wfs = new FileStream(unzipFileName, FileMode.Create, FileAccess.Write, FileShare.ReadWrite))
                            {
                                byte[] buff = new byte[2048];
                                int size = 0;
                                while (true)
                                {
                                    size = zis.Read(buff, 0, buff.Length);
                                    if (size > 0)
                                        wfs.Write(buff, 0, size);
                                    else
                                        break;
                                }
                                wfs.Flush();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("解压文件", ex);
            }
        }

        /// <summary>
        /// 日志
        /// </summary>
        /// <param name="message">信息</param>
        /// <param name="saveFolder">日志存放目录，文件名为log.txt</param>
        public static void Log(string message, string saveFolder)
        {
            try
            {
                if (!Directory.Exists(saveFolder))
                    Directory.CreateDirectory(saveFolder);

                using (StreamWriter sw = new StreamWriter(Path.Combine(saveFolder, "Log.txt"), true, Encoding.Default))
                {
                    sw.WriteLine(DateTime.Now.ToString() + " " + message);
                }
            }
            catch { }
        }

        /// <summary>
        /// 日志
        /// </summary>
        /// <param name="message">信息</param>
        /// <param name="logFileName">日志文件绝对路径</param>
        /// <param name="append">是否追加信息</param>
        public static void Log(string message, string logFileName, bool append)
        {
            try
            {
                if (!Directory.Exists(Path.GetDirectoryName(logFileName)))
                    Directory.CreateDirectory(Path.GetDirectoryName(logFileName));

                using (StreamWriter sw = new StreamWriter(logFileName, append, Encoding.Default))
                {
                    sw.WriteLine(DateTime.Now.ToString() + " " + message);
                }
            }
            catch { }
        }

        public string Resize(Match m)
        {

            try
            {
                return string.Format("font-size:{0}pt", (Convert.ToInt32(m.Groups[1].Value) * this._zoomPercent/100).ToString());
            }
            catch (Exception)
            {

                return m.Value;
            }
            

        }

        private string UpdateTagWidth(Match m)
        {

            try
            {
                if (_firstTag)
                {
                    _firstTag = false;
                    return string.Format("<table cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse: collapse; margin-left: 0pt;width:{0}\">",_tagWidth);
                }
                else
                {
                    return m.Value;
                }
            }
            catch (Exception)
            {

                return m.Value;
            }


        }



        /// <summary>
        /// 替换字符串
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="option"></param>
        /// <param name="ms"></param>
        /// <param name="rep"></param>
        /// <returns></returns>
        public static string ReplaceMatch(string expression, RegexOptions option, string ms, string rep)
        {
            Regex regex = new Regex(expression, option);
            string result = regex.Replace(ms, rep);
            return result;
        }

        /// <summary>
        /// 改变字体大小
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="ms"></param>
        /// <returns></returns>
        public string ChangeFontSizeMatch(string expression, string ms, int zoomPercent)
        {
            this._zoomPercent = zoomPercent;
            Regex regex = new Regex(expression, RegexOptions.IgnoreCase);
            MatchEvaluator me = new MatchEvaluator(Resize);
            string result = regex.Replace(ms, me);
            return result;
        }

        /// <summary>
        /// 替换第一个表格的宽度
        /// </summary>
        /// <param name="expression"></param>
        /// <param name="ms"></param>
        /// <returns></returns>
        public string UpdateTableWidthMatch(string expression, string ms, string width)
        {
            _firstTag = true;
            this._tagWidth = width;
            Regex regex = new Regex(expression, RegexOptions.IgnoreCase);
            MatchEvaluator me = new MatchEvaluator(UpdateTagWidth);
            return regex.Replace(ms, me);
        }

        /// <summary>
        /// 下载文件
        /// </summary>
        /// <param name="url"></param>
        /// <param name="saveName"></param>
        public static void DownloadToFile(string url, string saveName,CookieContainer cookie)
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            request.Credentials = CredentialCache.DefaultCredentials;
            if (cookie != null)
                request.CookieContainer = cookie;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();

            using (Stream dataStream = response.GetResponseStream())
            {
                using (FileStream fs = new FileStream(saveName, FileMode.Create, FileAccess.Write))
                {
                    byte[] buffer = new byte[1024];
                    while (true)
                    {
                        int sizeRead = dataStream.Read(buffer, 0, buffer.Length);
                        if (sizeRead > 0)
                            fs.Write(buffer, 0, sizeRead);
                        else
                            break;
                    }
                    fs.Flush();
                }
            }
        }

        /// <summary>
        /// 下载文件，并返回文件流
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static Stream DownloadToStream(string url)
        {
            WebRequest request = WebRequest.Create(url);
            request.Credentials = CredentialCache.DefaultCredentials;
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            return response.GetResponseStream();
        }

        /// <summary>
        /// 复制目录下的文件和子目录到另一个目录中
        /// </summary>
        /// <param name="src">源目录</param>
        /// <param name="dest">目标目录</param>
        public static void CopyDirectory(string src, string dest)
        {
            try
            {
                DirectoryInfo dir = new DirectoryInfo(src);
                foreach (FileSystemInfo fsi in dir.GetFileSystemInfos())
                {
                    string destName = Path.Combine(dest, fsi.Name);
                    if (fsi is FileInfo)
                    {
                        if (File.Exists(destName))
                            File.Delete(destName);
                        File.Copy(fsi.FullName, destName, true);
                    }
                    else
                    {
                        if (!Directory.Exists(destName))
                            Directory.CreateDirectory(destName);
                        CopyDirectory(fsi.FullName, destName);
                    }
                }
            }
            catch { }
        }

       

        /// <summary>
        /// 压缩目录
        /// </summary>
        /// <param name="FolderToZip">待压缩的文件夹，全路径格式</param>
        /// <param name="ZipedFile">压缩后的文件名，全路径格式</param>
        /// <returns></returns>
        public static bool ZipFolder(string FolderToZip, string ZipedFile)
        {
            bool res;
            if (!Directory.Exists(FolderToZip))
            {
                return false;
            }

            ZipOutputStream s = new ZipOutputStream(File.Create(ZipedFile));
            s.SetLevel(6);
            //s.Password = Password;

            res = ZipFileDictory(FolderToZip, s, "");

            s.Finish();
            s.Close();

            return res;
        }

        /// <summary>
        /// 递归压缩文件夹方法
        /// </summary>
        /// <param name="FolderToZip"></param>
        /// <param name="s"></param>
        /// <param name="ParentFolderName"></param>
        private static bool ZipFileDictory(string FolderToZip, ZipOutputStream s, string ParentFolderName)
        {
            bool res = true;
            string[] folders, filenames;
            ZipEntry entry = null;
            FileStream fs = null;
            Crc32 crc = new Crc32();

            try
            {

                //创建当前文件夹
                entry = new ZipEntry(Path.Combine(ParentFolderName, Path.GetFileName(FolderToZip) + "/"));  //加上 “/” 才会当成是文件夹创建
                s.PutNextEntry(entry);
                s.Flush();


                //先压缩文件，再递归压缩文件夹 
                filenames = Directory.GetFiles(FolderToZip);
                foreach (string file in filenames)
                {
                    //打开压缩文件
                    fs = File.OpenRead(file);

                    byte[] buffer = new byte[fs.Length];
                    fs.Read(buffer, 0, buffer.Length);
                    entry = new ZipEntry(Path.Combine(ParentFolderName, Path.GetFileName(FolderToZip) + "/" + Path.GetFileName(file)));

                    entry.DateTime = DateTime.Now;
                    entry.Size = fs.Length;
                    fs.Close();

                    crc.Reset();
                    crc.Update(buffer);

                    entry.Crc = crc.Value;

                    s.PutNextEntry(entry);

                    s.Write(buffer, 0, buffer.Length);
                }
            }
            catch
            {
                res = false;
            }
            finally
            {
                if (fs != null)
                {
                    fs.Close();
                    fs = null;
                }
                if (entry != null)
                {
                    entry = null;
                }
                GC.Collect();
            }


            folders = Directory.GetDirectories(FolderToZip);
            foreach (string folder in folders)
            {
                if (!ZipFileDictory(folder, s, Path.Combine(ParentFolderName, Path.GetFileName(FolderToZip))))
                {
                    return false;
                }
            }

            return res;
        }

        /// <summary>
        /// UTF8转换成GB2312
        /// </summary>
        /// <param name="utf8"></param>
        /// <returns></returns>
        public static string UTF82GB2312(string utf8)
        {
            Byte[] gb1 = System.Text.Encoding.UTF8.GetBytes(utf8);
            Byte[] gb2 = System.Text.Encoding.Convert(System.Text.Encoding.UTF8, System.Text.Encoding.GetEncoding("gb2312"), gb1);
            return System.Text.Encoding.GetEncoding("gb2312").GetString(gb2);

        }

        
        /// <summary>
        /// Tif转换至Jpg
        /// </summary>
        /// <param name="tifFilePath">tif文件路径</param>
        /// <param name="imgQuality"></param>
        /// <param name="combineImages">合并图片标识</param>
        /// <returns>Jpeg文件路径</returns>
        public static string Tif2Jpeg(string tifFilePath, int imgQuality,Boolean combineImages)
        {
            try
            {
                int len = tifFilePath.LastIndexOf(".tif");
                string fileName2 = tifFilePath.Substring(0, len);
                string filePath = fileName2.Substring(0, fileName2.LastIndexOf('\\') + 1);
                FileStream stream = File.OpenRead(tifFilePath);
                Bitmap bmp = new Bitmap(stream);
                System.Drawing.Image image = bmp;
                Guid objGuid = image.FrameDimensionsList[0];
                FrameDimension objDimension = new FrameDimension(objGuid);
                int totFrame = image.GetFrameCount(objDimension);
                int count = totFrame;

                if (count <= 0)
                {
                    bmp.Dispose();
                    return "";

                }

                Bitmap resultImg = null;
                Graphics resultGraphics = null;
                int imageHeight;
                for (int i = 0; i < totFrame; i++)//循环生成多张图片
                {
                    image.SelectActiveFrame(objDimension, i);

                    //合并图片
                    if (combineImages)
                    {
                        if (resultImg == null)
                        {
                            imageHeight = image.Height;

                            //resultImg = new Bitmap(tempImage.Width, tempImage.Height * imgCount);
                            resultImg = new Bitmap(792, 1120 * totFrame);
                            resultGraphics = Graphics.FromImage(resultImg);

                        }

                        if (i == 0)
                        {
                            resultGraphics.DrawImage(image, 0, 0);
                        }
                        else
                        {
                            resultGraphics.DrawImage(image, 0, 1120 * i);
                            //resultGraphics.DrawImage(tempImage, 0, imageHeight * i);
                        }
                    }

                    //保存图片
                    ImageUtility.CompressAsJPG(new Bitmap(image), filePath + "\\" + i + ".jpg", imgQuality);
                    //image.Save(filePath + "\\" + i + ".jpg", ImageFormat.Jpeg);
                }

                string imgFilePath = string.Empty;

                if (combineImages)
                {
                    imgFilePath = tifFilePath.Replace(".tif", ".jpg");
                    //ImageUtility.ThumbAsJPG(resultImg, result, (int)(resultImg.Width * 0.4), (int)(resultImg.Height * 0.4), quality);
                    ImageUtility.CompressAsJPG(resultImg, imgFilePath, imgQuality);
                    resultGraphics.Dispose();
                }
                else
                {
                    imgFilePath = filePath + "\\0.jpg";
                }
                bmp.Dispose();
                image.Dispose();
                stream.Close();
                //File.Delete(tifFilePath);
                return imgFilePath;
            }
            catch (Exception e)
            {
                return e.Message;
            }
        }


        /// <summary>
        /// 解压功能(解压压缩文件到指定目录)
        /// </summary>
        /// <param name="FileToUpZip">待解压的文件</param>
        /// <param name="ZipedFolder">指定解压目标目录</param>
        public static void UnZipFolder(string FileToUpZip, string ZipedFolder)
        {
            if (!File.Exists(FileToUpZip))
            {
                return;
            }

            if (!Directory.Exists(ZipedFolder))
            {
                Directory.CreateDirectory(ZipedFolder);
            }

            ZipInputStream s = null;
            ZipEntry theEntry = null;

            string fileName;
            FileStream streamWriter = null;
            try
            {
                s = new ZipInputStream(File.OpenRead(FileToUpZip));
                //s.Password = Password;
                while ((theEntry = s.GetNextEntry()) != null)
                {
                    if (theEntry.Name != String.Empty)
                    {
                        fileName = Path.Combine(ZipedFolder, theEntry.Name);
                        /**/
                        ///判断文件路径是否是文件夹
                        if (fileName.EndsWith("/") || fileName.EndsWith("\\"))
                        {
                            Directory.CreateDirectory(fileName);
                            continue;
                        }

                        streamWriter = File.Create(fileName);
                        int size = 2048;
                        byte[] data = new byte[2048];
                        while (true)
                        {
                            size = s.Read(data, 0, data.Length);
                            if (size > 0)
                            {
                                streamWriter.Write(data, 0, size);
                            }
                            else
                            {
                                break;
                            }
                        }
                    }
                }
            }
            finally
            {
                if (streamWriter != null)
                {
                    streamWriter.Close();
                    streamWriter = null;
                }
                if (theEntry != null)
                {
                    theEntry = null;
                }
                if (s != null)
                {
                    s.Close();
                    s = null;
                }
                GC.Collect();
            }
        }
        /// <summary>
        /// 转换字节大小，输出Byte、K、M、G、T
        /// </summary>
        /// <param name="bytesLength"></param>
        /// <returns></returns>
        public static string ToSize(int bytesLength)
        {
            if (bytesLength > 1024 * 1024 * 1024)
            {
                double newSize = Convert.ToDouble(bytesLength) / (1024 * 1024 * 1024);
                return newSize.ToString("0") + "G";
            }
            else if (bytesLength > 1024 * 1024)
            {
                double newSize = Convert.ToDouble(bytesLength) / (1024 * 1024);
                return newSize.ToString("0") + "M";
            }
            else if (bytesLength > 1024)
            {
                double newSize = Convert.ToDouble(bytesLength) / 1024;
                return newSize.ToString("0") + "K";
            }
            else
            {
                return bytesLength.ToString();
            }
        }

        /// <summary>
        /// 转换字节大小，输出Byte、K、M、G、T
        /// </summary>
        /// <param name="bytesLength"></param>
        /// <returns></returns>
        public static string ToSize(int bytesLength, string format)
        {
            if (bytesLength > 1024 * 1024 * 1024)
            {
                double newSize = Convert.ToDouble(bytesLength) / (1024 * 1024 * 1024);
                return newSize.ToString(format) + "G";
            }
            else if (bytesLength > 1024 * 1024)
            {
                double newSize = Convert.ToDouble(bytesLength) / (1024 * 1024);
                return newSize.ToString(format) + "M";
            }
            else if (bytesLength > 1024)
            {
                double newSize = Convert.ToDouble(bytesLength) / 1024;
                return newSize.ToString(format) + "K";
            }
            else
            {
                return bytesLength.ToString(format);
            }
        }

        /// <summary>
        /// 验证是否为邮箱地址
        /// </summary>
        /// <param name="inputEmail"></param>
        /// <returns></returns>
        public static bool IsEmail(string emailAddress)
        {
            string strRegex = @"\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*";
            Regex reg = new Regex(strRegex);
            if (reg.IsMatch(emailAddress))
                return (true);
            else
                return (false);
        }

        /// <summary>
        /// 验证是否为邮箱地址
        /// </summary>
        /// <param name="emailAddresses"></param>
        /// <param name="splitChar"></param>
        /// <returns></returns>
        public static bool IsEmail(string emailAddresses, char splitChar)
        {
            string[] emailArray = emailAddresses.Split(splitChar);
            foreach (string email in emailArray)
            {
                if (email.Trim() != string.Empty && !IsEmail(email))
                    return false;
            }
            return true;
        }

        /// <summary>
        /// 截取字符串过长的部分
        /// </summary>
        /// <param name="str">源字符串</param>
        /// <param name="length">需要保留的长度</param>
        /// <param name="endStr">最后结束字符</param>
        /// <returns></returns>
        public static string SubString(string str, int length, string endStr)
        {
            if (str == null || str.Length == 0)
                return str;

            if (str.Length <= length)
                return str;
            else
                return str.Substring(0, length - 1) + endStr;
        }

        /// <summary>
        /// 返回MD5
        /// </summary>
        /// <param name="ConvertString"></param>
        /// <returns></returns>
        public static string GetMd5Str(string ConvertString)
        {
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            string t2 = BitConverter.ToString(md5.ComputeHash(UTF8Encoding.Default.GetBytes(ConvertString)), 4, 8);
            t2 = t2.Replace("-", "");
            return t2;
        }

        /// <summary>
        /// 去除HTML标签
        /// </summary>
        /// <param name="strHtml"></param>
        /// <returns></returns>
        public static string NoHTML(string strHtml)
        {
            //删除脚本
            strHtml = Regex.Replace(strHtml, @"<script[^>]*?>.*?</script>", "",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"<\s*script.*?>[\s\S]*?<\s*/\s*script\s*>", "",
              RegexOptions.IgnoreCase);
            //删除样式
            strHtml = Regex.Replace(strHtml, @"<\s*style[^>]*>[^<>]*?<\s*/\s*style\s*>", "",
              RegexOptions.IgnoreCase);

            //删除HTML
            strHtml = Regex.Replace(strHtml, @"<(.[^>]*)>", "",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"([\r\n])[\s]+", "",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"-->", "", RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"<!--.*", "", RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(quot|#34);", "\"",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(amp|#38);", "&",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(lt|#60);", "<",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(gt|#62);", ">",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(nbsp|#160);", "   ",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(iexcl|#161);", "\xa1",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(cent|#162);", "\xa2",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(pound|#163);", "\xa3",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&(copy|#169);", "\xa9",
              RegexOptions.IgnoreCase);
            strHtml = Regex.Replace(strHtml, @"&#(\d+);", "",
              RegexOptions.IgnoreCase);

            strHtml.Replace("<", "");
            strHtml.Replace(">", "");
            strHtml.Replace("\r\n", "");
            //strHtml = HttpContext.Current.Server.HtmlEncode(strHtml).Trim();

            return strHtml;
        }


        /// <summary>
        /// 快速排序算法
        /// </summary>
        /// 快速排序为不稳定排序,时间复杂度O(nlog2n),为同数量级中最快的排序方法
        /// <param name="arr">划分的数组</param>
        /// <param name="low">数组低端上标</param>
        /// <param name="high">数组高端下标</param>
        /// <returns></returns>
        static int Partition(FileInfo[] arr, int low, int high)
        {
            //进行一趟快速排序,返回中心轴记录位置
            // arr[0] = arr[low];
            FileInfo pivot = arr[low];//把中心轴置于arr[0]
            while (low < high)
            {
                while (low < high && arr[high].CreationTime <= pivot.CreationTime)
                    --high;
                //将比中心轴记录小的移到低端
                Swap(ref arr[high], ref arr[low]);
                while (low < high && arr[low].CreationTime >= pivot.CreationTime)
                    ++low;
                Swap(ref arr[high], ref arr[low]);
                //将比中心轴记录大的移到高端
            }
            arr[low] = pivot; //中心轴移到正确位置
            return low;  //返回中心轴位置
        }
        /// <summary>
        /// Swaps the specified i.
        /// </summary>
        /// <param name="i">The i.</param>
        /// <param name="j">The j.</param>
        static void Swap(ref FileInfo i, ref FileInfo j)
        {
            FileInfo t;
            t = i;
            i = j;
            j = t;
        }

        /// <summary>
        /// 快速排序算法
        /// </summary>
        /// 快速排序为不稳定排序,时间复杂度O(nlog2n),为同数量级中最快的排序方法
        /// <param name="arr">划分的数组</param>
        /// <param name="low">数组低端上标</param>
        /// <param name="high">数组高端下标</param>
        public static void QuickSort(FileInfo[] arr, int low, int high)
        {
            if (low <= high - 1)//当 arr[low,high]为空或只一个记录无需排序
            {
                int pivot = Partition(arr, low, high);
                QuickSort(arr, low, pivot - 1);
                QuickSort(arr, pivot + 1, high);

            }
        }
        /// <summary>
        /// 获取文件(目录)的大小,单位:MB
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static long FileSize(string filePath)
        {
            long temp = 0;

            //判断当前路径所指向的是否为文件
            if (File.Exists(filePath) == false)
            {
                string[] str1 = Directory.GetFileSystemEntries(filePath);
                foreach (string s1 in str1)
                {
                    temp += FileSize(s1);
                }
            }
            else
            {

                //定义一个FileInfo对象,使之与filePath所指向的文件向关联,

                //以获取其大小
                FileInfo fileInfo = new FileInfo(filePath);
                return fileInfo.Length;
            }
            
            return temp / (1024*1024);
        }

        /// <summary>
        /// 清理目录
        /// </summary>
        /// <param name="dir">目录</param>
        /// <param name="maxFileCount">最大文件数</param>
        public static void ClearupDir(string dir, int maxFileCount)
        {

            try
            {

                FileInfo[] files = new DirectoryInfo(dir).GetFiles();
                var fileCount = files.Count();
                if (fileCount> maxFileCount)
                {
                    QuickSort(files, 0, files.Length - 1);//按时间排序

                    //删除500个旧的文件
                    int beginIndex = fileCount - 500;
                    beginIndex = (beginIndex >= 0) ? beginIndex : 0;
                    for (int i=beginIndex; i < fileCount; i++)
                    {
                        //删除对应的目录
                        var tempDir = dir + Path.GetFileNameWithoutExtension(files[i].FullName);
                        RemoteDelDir(tempDir);
                        //删除文件
                        files[i].Delete();
                     }


                }
            }
            catch (Exception)
            {

                
            }
           
        }
        /// <summary>
        /// 路径末端自动加上斜杠
        /// </summary>
        /// <param name="path">路径</param>
        /// <returns></returns>
        public static string PathIncludeSlash(string path)
        {
            if (path.EndsWith(@"\") || path.EndsWith("/"))
            {
                return path;
            }
            else
            {
                return path + @"\";
            }
        }
        /// <summary>
        /// 获取本机网卡的MAC列表
        /// </summary>
        /// <returns>MAC列表</returns>
        public static string[] GetLocalMACAddresses()
        {
            string[] macAddresses = { string.Empty, string.Empty, string.Empty, string.Empty};
            NetworkInterface[] nis = NetworkInterface.GetAllNetworkInterfaces();
            var i = 0;
            foreach (NetworkInterface ni in nis)
            {
                PhysicalAddress pa = ni.GetPhysicalAddress();
                macAddresses[i++] = pa.ToString();
                if (i >= 4)
                    break;
            }

            return macAddresses;
           
        }

        /// <summary>
        /// 删除目录
        /// </summary>
        /// <param name="dir">目录</param>
        /// <returns></returns>
        public static bool RemoteDelDir(string dir)
        {
            FileInfo fileInfo = new FileInfo(dir);
            if (fileInfo.Exists)
            {
                return SigleFileDel(dir);
            }
            else
            {
                DirectoryInfo Dir = new DirectoryInfo(dir);
                if (!Dir.Exists) return false;
                if (Dir.GetFiles().Length <= 0)
                {
                    Dir.Delete();
                    return true;
                }
                else
                {
                    foreach (FileInfo file in Dir.GetFiles())
                    {
                        string FilePath = file.FullName;
                        if (!SigleFileDel(FilePath))
                            return false;
                    }
                    Dir.Delete();
                    return true;
                }
            }
        }
        /// <summary>
        /// 删除文件
        /// </summary>
        /// <param name="filePath">文件路径</param>
        /// <returns></returns>
        public static bool SigleFileDel(string filePath)
        {
            FileInfo file = new FileInfo(filePath);
            if (file.Exists)
            {
                file.Delete();
                return true;
            }
            else
            {
                return false;
            }
        }




    }

}

