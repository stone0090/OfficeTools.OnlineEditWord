using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Aspose.Words;
using System.Web;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Checksums;
using Aspose.Words.Fields;
using System.Net;

namespace LenovoCW.MOA
{   
        public delegate void ConvertChangeEventHandler(string fileName);
        public delegate void ConvertStausEventHandler(bool status);

        /// <summary>
        /// 转换引擎类
        /// </summary>
        public class ConvertEngine
        {

            public event ConvertChangeEventHandler ConvertPage;       //页面转换
            public event ConvertStausEventHandler ConvertFinish;     //转换完成

            /// <summary>
            /// 临时文件目录的最大文件数量
            /// </summary>
            public static int SaveDirMaxFileCount = 3000;

            private string _tempPath=string.Empty;
            private string _relativeTempPath = string.Empty;
            private string _fileServer = string.Empty;
            private string _serverUser = string.Empty;
            private string _serverPwd = string.Empty;

            private CookieContainer _cookie = null;

            private Boolean _combineImages;

            /// <summary>
            /// 许可的MAC地址列表
            /// </summary>
            private string[] _macAddresses = { string.Empty, string.Empty, string.Empty, string.Empty };
            /// <summary>
            /// 方正原版阅读器应用程序的完整路径
            /// </summary>
            //public string CebAppPath
            //{
            //    get;
            //    set;
            //}
            /// <summary>
            /// 虚拟打印机打印文件保存路径
            /// </summary>
            public string CebImageSavePath
            {
                get;
                set;
            }

            /// <summary>
            /// 当临时文件超过一定数量时,自动删除旧文件的标识
            /// </summary>
            public Boolean AutoDelOldFile
            {
                get;
                set;
            }

            /// <summary>
            /// 设置许可
            /// </summary>
            /// <param name="macAddress">MAC地址</param>
            private void SetLicense(string macAddress)
            {
                this._macAddresses[0] = macAddress;
            }

            /// <summary>
            /// 解密字符串
            /// </summary>
            /// <param name="input">源字符串</param>
            /// <returns>解密后的字符串</returns>
            private string DecryptString(string input)
            {
                if (input.Equals(string.Empty))
                {
                    return input;
                }
                try
                {

                    byte[] byKey = { 0x33, 0xE8, 0x65, 0x3E, 0x49, 0x75, 0x61, 0x6E };
                    byte[] IV = { 0xFE, 0xDC, 0xBA, 0x48, 0x76, 0x34, 0x32, 0x10 };
                    byte[] inputByteArray = new Byte[input.Length];
                    DESCryptoServiceProvider des = new DESCryptoServiceProvider();
                    inputByteArray = System.Convert.FromBase64String(input);
                    MemoryStream ms = new MemoryStream();
                    CryptoStream cs = new CryptoStream(ms, des.CreateDecryptor(byKey, IV), CryptoStreamMode.Write);
                    cs.Write(inputByteArray, 0, inputByteArray.Length);
                    cs.FlushFinalBlock();
                    Encoding encoding = new UTF8Encoding();
                    return encoding.GetString(ms.ToArray());
                }
                catch (Exception)
                {

                    return input;
                }
            }

            /// <summary>
            /// 许可检验
            /// </summary>
            /// <returns></returns>
            private Boolean CheckLicense()
            {
                return true;
               
            }


            /// <summary>
            /// 构造函数
            /// </summary>
            /// <param name="tempPath">临时文件存放路径(注意别漏了最后的\字符),如:d:\temp\</param>
            /// <param name="relativeTempPath">tempPath参数对应的相对路径(注意别漏了最后的\字符),如:\temp\</param>
            /// <param name="serverIP">文件源存放的服务器IP地址,如通过HTTP协议的方式,可设置为nil</param>
            /// <param name="serverLoginName">文件源存放的服务器访问帐户,如通过HTTP协议的方式,可设置为nil</param>
            /// <param name="serverLoginPassword">文件源存放的服务器访问密码,如通过HTTP协议的方式,可设置为nil</param>
            public ConvertEngine(string tempPath,string relativeTempPath,string serverIP,string serverLoginName,string serverLoginPassword)
            {

                this._tempPath = Utils.PathIncludeSlash(tempPath);
                this._relativeTempPath = Utils.PathIncludeSlash(relativeTempPath);
                this._fileServer = serverIP;
                this._serverUser = serverLoginName;
                this._serverPwd = serverLoginPassword;

                this.AutoDelOldFile = false;

                this._macAddresses = Utils.GetLocalMACAddresses();

                try
                {
                    if (!Directory.Exists(this._tempPath))
                    {
                        Directory.CreateDirectory(this._tempPath);
                    }
                }
                catch (Exception e)
                {
                    
                    throw e;
                }

            }
            /// <summary>
            /// 设置下载链接使用的Cookie
            /// </summary>
            /// <param name="cookie"></param>
            public void SetCookie(CookieContainer cookie)
            {
                _cookie = cookie;
            }



            /// <summary>
            /// 打印日志到文件
            /// </summary>
            /// <param name="log">日志内容</param>
            private  void WriteLog(string log)
            {
                if (!Directory.Exists(_tempPath))
                {
                    Directory.CreateDirectory(_tempPath);
                }
                using (StreamWriter sw = new StreamWriter(_tempPath + "log.txt", true))
                {
                    sw.WriteLine(System.DateTime.Now.ToString().Trim() + "  :  " + log);
                    sw.Flush();
                    sw.Close();
                    sw.Dispose();
                }
            }

            /// <summary>
            /// 转换文档
            /// </summary>
            /// <param name="srcPath">原文件路径</param>
            /// <param name="destFileExtension">目标类型,如:.html</param>
            /// <param name="srcIsZip">源文件是否是压缩文件</param>
            /// <param name="pageIndex">页码</param>
            /// <param name="imgQuality">图片质量</param>
            /// <param name="htmlEncoding">转换的编码</param>
            /// <returns>如果是HTML,返回的是HTML的内容,否则返回转换后保存的路径</returns>
            ///
            [STAThread]
            private string _Convert(string srcPath, string destFileExtension, Boolean srcIsZip, int pageIndex = -1, int imgQuality = 0, string htmlEncoding = "gb2312", int fontZoomPercent=100,Boolean combineImages=false)
            {
                //判断许可
                //if (!CheckLicense())
                //{
                //    return new InternalBufferOverflowException().ToString();
                //}


                //清理旧文件
                if (this.AutoDelOldFile)
                {
                    Utils.ClearupDir(this._tempPath, SaveDirMaxFileCount);

                }

                //文件名唯一标识
                var fileGUID = Utils.GetMd5Str(srcPath);

                try
                {
                    this._combineImages = combineImages;

                    string fileExtension = Path.GetExtension(srcPath).ToLower();
                    string fileName = fileGUID + fileExtension;

                    string tempFilePath = this._tempPath + fileName;
                    //判断目录是否存在（如果存在则说明转换过），如果是则不在下载，而doc每次都更新
                    //var dirInfo = new DirectoryInfo(this._tempPath + fileGUID);
                    //if (!dirInfo.Exists || dirInfo.GetFiles().Length == 0 || fileExtension == ".doc")
                    {
                        //如果是HTTP链接，则通过HTTP下载
                        if (srcPath.Contains("http://"))
                        {
                            var downloadTempFile = tempFilePath + ".tmp";
                            Utils.DownloadToFile(srcPath, downloadTempFile,_cookie);
                            //如果是压缩文件,则解压
                            if (srcIsZip)
                            {
                                Utils.UnZipFile(downloadTempFile, tempFilePath);
                            }
                            else
                            {
                                File.Copy(downloadTempFile,tempFilePath);
                            }
                            FileInfo fi = new FileInfo(downloadTempFile);
                            fi.Delete();
                        }
                        //如果是局域网，则直接获取
                        else if (!string.IsNullOrEmpty(this._fileServer) && srcPath.StartsWith(@"\\"))
                        {

                            using (IdentityScope identity = new IdentityScope(this._serverUser, this._fileServer, this._serverPwd))
                            {
                                //如果是压缩文件,则解压
                                if (srcIsZip)
                                {
                                    Utils.UnZipFile(srcPath, tempFilePath);
                                }
                                else
                                {
                                    tempFilePath = srcPath;
                                }

                            }

                        }
                        else
                        {
                            //如果是压缩文件,则解压
                            if (srcIsZip)
                            {
                                Utils.UnZipFile(srcPath, tempFilePath);
                            }
                            else
                            {
                                tempFilePath = srcPath;
                            }
                        }



                        ////如果是压缩文件,则解压
                        //if (srcIsZip)
                        //{
                        //    try
                        //    {
                        //        //如果是局域网的，则先登录认证
                        //        if (!string.IsNullOrEmpty(this._fileServer) && srcPath.StartsWith(@"\\"))
                        //        {
                        //            using (IdentityScope identity = new IdentityScope(this._serverUser, this._fileServer, this._serverPwd))
                        //            {
                        //                Utils.UnZipFile(srcPath, tempFilePath);
                        //            }
                        //        }
                        //        else
                        //        {
                        //            Utils.UnZipFile(srcPath, tempFilePath);
                        //        }

                        //    }
                        //    catch (Exception)
                        //    {
                        //        return "文件解压失败,请确认源文件是压缩文件.";
                        //    }
                        //}
    
                    }
                     

                    //如果路径为空，则直接返回
                    if (string.IsNullOrEmpty(tempFilePath))
                        return "";

                    string destFileName = fileGUID + destFileExtension;
                    string destDir = this._tempPath + fileGUID + "\\";
                    string destFilePath = destDir + destFileName;
                    string resultData = string.Empty;
 
                    if (fileExtension == ".doc" || fileExtension == ".docx")
                    {
                        if (!Directory.Exists(destDir))
                            Directory.CreateDirectory(destDir);

                        //转换成HTML
                        if (destFileExtension.Equals(".html"))
                        {
                            Document doc = new Document(tempFilePath);

                            //不显示文本边框
                            doc.SaveOptions.HtmlExportTextInputFormFieldAsText = true;
                            doc.SaveOptions.HtmlExportAllowNegativeLeftIndent = true;
                            doc.SaveOptions.HtmlExportDocumentProperties = false;
                            doc.AcceptAllRevisions();


                            FormFieldCollection formFields = doc.Range.FormFields;
                            for (int i =  formFields.Count-1; i >= 0; i--)
                            {
                                var ff = formFields[i];
                                if (ff.Type == FieldType.FieldNone)
                                {
                                    ff.Remove();
                                }

                            }

                            doc.Save(destFilePath);


                            StreamReader sr = new StreamReader(destFilePath);
                            var htmlData = Utils.ReplaceMatch("<meta name=\"generator\" content=\".*?\" />", RegexOptions.None, sr.ReadToEnd(), "");
                            htmlData = Utils.ReplaceMatch(@"<html>[\s\S]*?<body>", RegexOptions.None, htmlData, "");
                            htmlData = Utils.ReplaceMatch(@"</body>[\s\S]*?</html>", RegexOptions.None, htmlData, "");
                            byte a = 0xc2;
                            byte b = 0xa0;
                            htmlData = htmlData.Replace((char)a, ' ');
                            htmlData = htmlData.Replace((char)b, ' ');
                            htmlData = Utils.ReplaceMatch("<img src=\"(.*?)\"", RegexOptions.IgnoreCase, htmlData, "<img src=\""+this._relativeTempPath + fileGUID+"/$1\"");
                            string pageData = "<span style=\"font-family:'宋体'; font-size:[1-9]+?pt\">－</span><span[^>].*?>[1-9]+?</span><span style=\"font-family:'宋体'; font-size:[1-9]+?pt\">－</span>";
                            htmlData = Utils.ReplaceMatch(pageData, RegexOptions.IgnoreCase, htmlData, "");
                            var utils = new Utils();
                            htmlData = utils.UpdateTableWidthMatch("<table cellspacing=\"0\" cellpadding=\"0\" style=\"border-collapse[^>]*?>", htmlData, "100%");
                            htmlData = htmlData.Replace("442.35pt", "100%");

                            sr.Close();
                            
                            if (htmlEncoding.ToLower().Equals("gb2312"))
                            {
                                //StreamWriter sw = new StreamWriter(destFilePath, false, Encoding.GetEncoding("gb2312"));
                                //sw.Write(htmlData);
                                //sw.Flush();
                                //sw.Close();

                                resultData = Utils.UTF82GB2312(htmlData);
                            }
                            else 
                            {
                                //StreamWriter sw = new StreamWriter(destFilePath);
                                //sw.Write(htmlData);
                                //sw.Flush();
                                //sw.Close();

                                resultData = htmlData;
                               
                            }

                            //删除临时文件
                            Utils.SigleFileDel(destFilePath);
                            //如果文件夹为空，则删除文件夹
                            var dir = new DirectoryInfo(destDir);
                            if (dir.Exists && dir.GetFiles().Length==0)
                                Utils.RemoteDelDir(destDir);
                        }
                        //转换成图片的
                        else if (destFileExtension.Equals(".jpg"))
                        {
                            resultData = new DocEngine().Convert(tempFilePath, this._tempPath, fileGUID,combineImages, pageIndex, imgQuality);
                           
                        }

                    }
                    else if (fileExtension == ".pdf")
                    {
                        resultData = new PDFEngine().Convert(tempFilePath, this._tempPath, fileGUID,combineImages, pageIndex, imgQuality);



                    }
                    else if (fileExtension == ".ceb")
                    {

                        if (string.IsNullOrEmpty(CebImageSavePath))
                            resultData = new CebEngine().Convert(tempFilePath, this._tempPath, fileGUID, this._tempPath, combineImages, imgQuality);
                        else
                            resultData = new CebEngine().Convert(tempFilePath, this._tempPath, fileGUID, CebImageSavePath, combineImages, imgQuality); ;
                                               
                    }

                    //删除临时文件
                    Utils.SigleFileDel(tempFilePath);
 
                    return resultData;
                }

                catch (Exception)
                {

                    return "未能找到文件。";
                }
            }

            /// <summary>
            /// 转换成HTML
            /// </summary>
            /// <param name="fileURL">源文件路径</param>
            /// <param name="isZip">文件是否是压缩包格式</param>
            /// <returns></returns>
            public string ConvertToHtml(string fileURL, Boolean isZip)
            {

                return ConvertToHtml(fileURL, isZip, "utf-8");
                //OnConvertFinish(true);
            }


            /// <summary>
            /// 转换成HTML
            /// </summary>
            /// <param name="fileURL">源文件路径,支持HTTP协议,局域网共享等方式</param>
            /// <param name="isZip">文件是否是压缩包格式</param>
            ///<param name="encoding">文档编码</param>
            ///<param name="zoomPercent">字体缩放比例</param>
            /// <returns></returns>
            public string ConvertToHtml(string fileURL, Boolean isZip, string encoding, int zoomPercent=100)
            {
                return this._Convert(fileURL, ".html", isZip, -1,100,encoding,zoomPercent);

            }

            /// <summary>
            /// 转换接口
            /// </summary>
            /// <param name="fileURL">源文件路径</param>
            /// <returns></returns>
            public string Convert(string fileURL)
            {
                var ext = Path.GetExtension(fileURL).ToLower();
                if (ext == ".pdf")
                    return this.Convert(fileURL,true,true,null,50);
                else if (ext == ".ceb")
                {
                    Boolean fileZiped = (Path.GetFileNameWithoutExtension(fileURL).Length > 20)?true:false;
                    return this.Convert(fileURL, true, fileZiped, null, 60);
                }
                else if (ext == ".doc")
                    return this.Convert(fileURL, true, true, "gb2312", 0,80);
                else
                    return this.Convert(fileURL, true, false, "gb2312", 0,80);
            }

            /// <summary>
            /// 转换接口
            /// </summary>
            /// <param name="fileURL">源文件路径</param>
            /// <param name="zoomPercent">缩放比例,取值：1--100</param>
            /// <returns></returns>
            public string Convert(string fileURL, int zoomPercent)
            {
                var ext = Path.GetExtension(fileURL).ToLower();
                if (ext == ".pdf")
                    return this.Convert(fileURL, true, true, null, 50);
                else if (ext == ".ceb")
                {
                    Boolean fileZiped = (Path.GetFileNameWithoutExtension(fileURL).Length > 20) ? true : false;
                    return this.Convert(fileURL, true, fileZiped, null, 60);
                }
                else if (ext == ".doc")
                    return this.Convert(fileURL, true, true, "gb2312", 0, zoomPercent);
                else
                    return this.Convert(fileURL, true, false, "gb2312", 0, zoomPercent);
            }

            /// <summary>
            /// 转换接口
            /// </summary>
            /// <param name="fileURL">源文件路径,支持HTTP协议,局域网共享等方式</param>
            /// <param name="containHtmlTag">返回的内容是否包含HTML标签标识</param>
            /// <param name="isZip">是否是压缩文件标识</param>
            /// <param name="docEncoding">文档编解码,默认为utf-8，取值:gb2312,utf-8</param>
            /// <param name="imgQuality">图片质量,默认50,取值范围:1-100</param>
            /// <returns></returns>
            public string Convert(string fileURL, Boolean containHtmlTag,Boolean isZip, string docEncoding, int imgQuality,int zoomPercent=100)
            {
                var ext = Path.GetExtension(fileURL).ToLower();
                if (ext == ".pdf" || ext == ".ceb")
                    return this.ConvertToImages(fileURL, imgQuality, isZip,true,false);
                else if (ext == ".doc")
                    return this.ConvertToHtml(fileURL, isZip, docEncoding, zoomPercent);
                else
                    return this.ConvertToHtml(fileURL, isZip, docEncoding, zoomPercent);
            }

            /// <summary>
            /// 转换所有页为图片
            /// </summary>
            /// <param name="fileURL">源文件路径</param>
            /// <param name="imgQuality">图片质量,取值范围:1-100</param>
            /// <param name="isZip">压缩文件</param>
            /// <param name="resultIncludePageCount">返回的结果中包含有页码</param>
            /// <returns>返回图片对应的目录,如果resultIncludePageCount为True，返回的格式为：图片目录|页码</returns>
            public string ConvertToImages(string fileURL, int imgQuality, Boolean isZip,Boolean combineImages,Boolean resultIncludePageCount)
            {
                try
                {
                    var path = this._Convert(fileURL, ".jpg", isZip, -1, imgQuality, "gb2312", 100, combineImages);
                    path = Path.GetDirectoryName(GetRelativePath(path))+"/";

                    if (resultIncludePageCount)
                        path =path +"|"+ GetPageCount(fileURL).ToString();
                    return path;
                }
                catch (Exception)
                {

                    return "";
                }
            }

            /// <summary>
            /// 转换成图片
            /// </summary>
            /// <param name="fileURL">源文件路径</param>
            /// <param name="pageIndex">页码</param>
            /// <returns>HTML描述的对应页码图片标签</returns>
            public string ConvertToImage(string fileURL, int pageIndex,int imgQuality)
            {
                try
                {

                    var ext = Path.GetExtension(fileURL).ToLower();
                    var path = string.Empty;
                    if (ext == ".pdf")
                    {
                        path = this._Convert(fileURL, ".jpg", true, -1, imgQuality);

                        
                    }
                    else if (ext == ".ceb")
                    {
                        Boolean fileZiped = (Path.GetFileNameWithoutExtension(fileURL).Length > 20) ? true : false;
                        path = this._Convert(fileURL, ".jpg", fileZiped, -1, imgQuality);
                    }
                    else if (ext == ".doc")
                    {
                        path = this._Convert(fileURL, ".jpg", true, -1, imgQuality);
                    }


                    //根据页码返回值
                    var pageJpgPath = string.Empty;
                    if (path.Length > 0)
                    {
                        pageIndex--;
                        pageJpgPath = Path.GetDirectoryName(path) + "\\" + pageIndex + ".jpg";

                        if (!File.Exists(pageJpgPath))
                        {
                            return path;
                        }
                    }

                    var resultPath = AddImgTag(GetRelativePath(pageJpgPath));


                    return resultPath;

                }
                catch (Exception e)
                {

                    return e.Message;
                }
            }
            /// <summary>
            /// 转换成图片
            /// </summary>
            /// <param name="fileURL">源文件路径</param>
            /// <returns></returns>
            public string ConvertToImage(string fileURL)
            {
               return ConvertToImage(fileURL, 50,false);
            }


            /// <summary>
            /// 转换成图片
            /// </summary>
            /// <param name="fileURL">源文件路径</param>
            /// <param name="imgQuality">图片质量</param>
            /// <returns>返回值格式：图片存放目录|页数</returns>
            public string ConvertToImage(string fileURL,int imgQuality)
            {
                return ConvertToImage(fileURL, imgQuality, false);
            }

            /// <summary>
            /// 转换成图片
            /// </summary>
            /// <param name="fileURL">源文件路径</param>
            /// <param name="imgQuality">图片质量</param>
            /// <param name="combineImages">合并图片的标识</param>
            /// <returns>返回值格式：图片存放目录|页数</returns>
            public string ConvertToImage(string fileURL, int imgQuality,Boolean combineImages)
            {
                try
                {
                    var ext = Path.GetExtension(fileURL).ToLower();
                    var path = string.Empty;
                    if (ext == ".pdf")
                    {
                        path = this._Convert(fileURL, ".jpg", true, -1, imgQuality,"gb2312", 100, combineImages);

                    }
                    else if (ext == ".ceb")
                    {
                        Boolean fileZiped = (Path.GetFileNameWithoutExtension(fileURL).Length > 20) ? true : false;
                        path = this._Convert(fileURL, ".jpg", fileZiped, -1, imgQuality, "gb2312", 100, combineImages);
                    }
                    else if (ext == ".doc")
                    {
                        path = this._Convert(fileURL, ".jpg", true, -1, imgQuality, "gb2312", 100, combineImages);
                    }

                    var resultPath = string.Empty;
                    if (path.Length > 0)
                    {
                        resultPath = Path.GetDirectoryName(GetRelativePath(path)) + "/";
                    }
                    else
                        resultPath = path;



                    var pageCount = GetPageCount(fileURL);
                    var result = string.Format("{0}|{1}", resultPath, pageCount);

                    return result;

                }
                catch (Exception e)
                {

                    return e.Message;
                }
            }

            /// <summary>
            /// 根据路径返回页码数
            /// </summary>
            /// <param name="fileURL"></param>
            /// <returns></returns>
            public int GetPageCount(string fileURL)
            {
                try
                {
                    var guid = Utils.GetMd5Str(fileURL);

                    var files = new DirectoryInfo(this._tempPath + guid).GetFiles();
                    var ext = Path.GetExtension(fileURL).ToLower();

                    int pageCount = 0;
                    if (this._combineImages)
                        pageCount = files.Length - 1;
                    else
                        pageCount = files.Length;
                        
                    
                    if (ext == ".ceb" || ext == ".doc")
                        return pageCount - 1;
                    else
                        return pageCount;

                }
                catch (Exception)
                {

                    return 0;
                }
            }
              

            

            /// <summary>
            /// 获取相对路径
            /// </summary>
            /// <param name="path"></param>
            /// <returns></returns>
            private string GetRelativePath(string path)
            {

                if (path.Length > 0 && this._relativeTempPath.Length > 0)
                {
                        string[] paths = { };
                        
                        if (path.Contains('/'))
                        {
                            paths = path.Split('/');
                            if (paths.Length < 2)
                                return path;

                            return this._relativeTempPath + paths[paths.Length - 2] + "/" + Path.GetFileName(path);
                        }
                        else
                        {
                            paths = path.Split('\\');
                            if (paths.Length < 2)
                                return path;

                            return  this._relativeTempPath + paths[paths.Length - 2] + @"\" + Path.GetFileName(path);
                        }

                }
                else
                    return path;
            }
            /// <summary>
            /// 添加IMG标签
            /// </summary>
            /// <param name="path"></param>
            /// <returns></returns>
            private string AddImgTag(string path)
            {
                if (path.Length > 0)
                {
                    return string.Format("<img src=\"{0}\" />", path);
                }
                else
                {
                    return "";
                }
            }

            private void OnConvertPage(string fileName)
            {
                if (ConvertPage != null)
                    ConvertPage(fileName);
            }

            private void OnConvertFinish(bool satus)
            {
                if (ConvertFinish != null)
                    ConvertFinish(satus);
            }

        }
}
