using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Net;
using System.Runtime.CompilerServices;
using Aspose.Words;
using Aspose.Words.Saving;
using CSN.DotNetLibrary.Compression.Zip;
using CSN.DotNetLibrary.Security.WindowsAuthentications;
using O2S.Components.PDFRender4NET;
using System.Threading;
using Microsoft.Office.Interop.Word;

namespace CSN.DotNetLibrary.OfficeHelper
{
    public class ConvertEngine
    {
        // variable
        private string _serverIP = string.Empty;
        private string _serverLogOnName = string.Empty;
        private string _serverLogOnPassword = string.Empty;
        private string _tempPath = string.Empty;
        private string _relativeTempPath = string.Empty;
        private int _saveDirMaxFileCount;
        private DateTime _deleteDateTime;

        [CompilerGenerated]
        private bool _autoDelOldFile;
        private delegate void ClearupDirDelegate(string directory, int dirMaxFileCount, DateTime deleteDateTime);

        // property
        public bool AutoDelOldFile
        {
            [CompilerGenerated]
            get { return _autoDelOldFile; }
            [CompilerGenerated]
            set { _autoDelOldFile = value; }
        }
        public int SaveDirMaxFileCount
        {
            get { return _saveDirMaxFileCount; }
            set
            {
                ValidateHelper.Begin().InRange(value, 1, 99999);
                _saveDirMaxFileCount = value;
            }
        }
        public DateTime DeleteDateTime
        {
            get { return _deleteDateTime; }
            set { _deleteDateTime = value; }
        }

        // constructor
        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="tempPath">临时文件存放路径(注意别漏了最后的\字符),如:d:\temp\</param>
        /// <param name="relativeTempPath">tempPath参数对应的相对路径(注意别漏了最后的\字符),如:\temp\</param>
        /// <param name="serverIP">文件源存放的服务器IP地址,如通过HTTP协议的方式,可设置为nil</param>
        /// <param name="serverLogOnName">文件源存放的服务器访问帐户,如通过HTTP协议的方式,可设置为nil</param>
        /// <param name="serverLogOnPassword">文件源存放的服务器访问密码,如通过HTTP协议的方式,可设置为nil</param>
        public ConvertEngine(string tempPath, string relativeTempPath, string serverIP, string serverLogOnName, string serverLogOnPassword)
        {
            this._tempPath = PathIncludeSlash(tempPath);
            this._relativeTempPath = PathIncludeSlash(relativeTempPath);
            this._serverIP = serverIP;
            this._serverLogOnName = serverLogOnName;
            this._serverLogOnPassword = serverLogOnPassword;
            this._saveDirMaxFileCount = 1000;
            this._deleteDateTime = DateTime.Parse(DateTime.Now.AddDays(-1).ToShortDateString() + " 23:59:59");
            if (!Directory.Exists(this._tempPath)) { Directory.CreateDirectory(this._tempPath); }
        }

        // public method
        /// <summary>
        /// PDF,Wrod转换成图片
        /// </summary>
        /// <param name="fileUrl">源文件路径</param>
        /// <returns>返回值格式：图片存放目录|页数</returns>
        public string ConvertToImage(string fileUrl)
        {
            return this.ConvertToImage(fileUrl, ImageFormat.Jpeg, 100, 50);
        }
        /// <summary>
        /// PDF,Wrod转换成图片
        /// </summary>
        /// <param name="fileUrl">源文件路径</param>
        /// <param name="imgQuality">图片质量</param>
        /// <returns>返回值格式：图片存放目录|页数</returns>
        public string ConvertToImage(string fileUrl, int imgQuality)
        {
            return this.ConvertToImage(fileUrl, ImageFormat.Jpeg, 100, imgQuality);
        }
        /// <summary>
        /// PDF,Wrod转换成图片
        /// </summary>
        /// <param name="fileUrl">源文件路径</param>
        /// <param name="imgResolution">图片分辨率</param>
        /// <param name="imgQuality">图片质量</param>
        /// <returns>返回值格式：图片存放目录|页数</returns>
        public string ConvertToImage(string fileUrl, int imgResolution, int imgQuality)
        {
            return this.ConvertToImage(fileUrl, ImageFormat.Jpeg, imgResolution, imgQuality);
        }
        /// <summary>
        /// PDF,Wrod转换成图片
        /// </summary>
        /// <param name="fileUrl">源文件路径</param>
        /// <param name="imgFormat">图片格式</param>
        /// <param name="imgResolution">图片分辨率</param>
        /// <param name="imgQuality">图片质量</param>
        /// <returns>返回值格式：图片存放目录|页数</returns>
        public string ConvertToImage2(string fileUrl, ImageFormat imgFormat, int imgResolution, int imgQuality)
        {
            DateTime ConvertStartTime = DateTime.Now;
            ConvertResult convertResult = ConvertResult.HasException;
            int pageCount = 0;
            string docId = string.Empty;
            string tempFilePath = string.Empty;
            string tempFilePath2 = string.Empty;
            string tempImageDirPath = string.Empty;
            string extension = Path.GetExtension(fileUrl).ToLower();
            string logFilePath = Path.Combine(this._tempPath, "Log_" + ConvertStartTime.ToString("yyyyMMdd")) + ".txt";
             
            try
            {
                // validate parameter
                ValidateHelper.Begin().NotNullAndEmpty(fileUrl).NotNull(imgFormat)
                    .InRange(imgResolution, 1, 512).InRange(imgQuality, 1, 100).CheckFileType(extension);

                docId = Path.GetFileNameWithoutExtension(fileUrl);
                tempFilePath = Path.Combine(this._tempPath, Guid.NewGuid().ToString()) + extension;
                tempFilePath2 = Path.Combine(this._tempPath, Guid.NewGuid().ToString()) + extension;

                // if _autoDelOldFile is ture , we need clearup old file by date and ignore exception
                if (this._autoDelOldFile)
                {
                    ClearupDirDelegate clearupDirDelegate = ClearupDir;
                    clearupDirDelegate(this._tempPath, this._saveDirMaxFileCount, this._deleteDateTime);
                }

                // three case for copy file to local temp directory , 1.from Http:// , 2.from other PC , 3.from local
                if (fileUrl.StartsWith("http://") || fileUrl.StartsWith("https://"))
                {
                    DownloadToFile(fileUrl, tempFilePath);
                }
                else if (fileUrl.StartsWith(@"\\"))
                {
                    using (WindowsIdentityScope wis = new WindowsIdentityScope(this._serverIP, this._serverLogOnName, this._serverLogOnPassword))
                    {
                        File.Copy(fileUrl, tempFilePath);
                    }
                }
                else
                {
                    File.Copy(fileUrl, tempFilePath);
                }

                // if the file is ziped ,we need decompress it
                if (ZipFile.IsFileZiped(tempFilePath))
                {
                    ZipFile.DecompressFirstFile(tempFilePath, tempFilePath2);
                }
                else
                {
                    tempFilePath2 = tempFilePath;
                }

                string tempImageDirName = docId + "&" + File.GetLastWriteTime(tempFilePath2).ToFileTime().ToString();
                string tempRelativeImageDirPath = PathIncludeSlash(Path.Combine(this._relativeTempPath, tempImageDirName));
                tempImageDirPath = PathIncludeSlash(Path.Combine(this._tempPath, tempImageDirName));

                if (Directory.Exists(tempImageDirPath))
                {
                    DirectoryInfo tempImageDirInfo = new DirectoryInfo(tempImageDirPath);
                    if (tempImageDirInfo.GetFiles().Length > 0)
                    {
                        convertResult = ConvertResult.ReturnCashe;
                        return tempRelativeImageDirPath + "|" + tempImageDirInfo.GetFiles().Length;
                    }
                }

                // convert file to image
                if (extension == ".docx" || extension == ".doc")
                {
                    pageCount = this.ConvertWord2Image(tempFilePath2, tempImageDirPath, string.Empty, 0, 0, imgFormat, imgResolution, imgQuality);
                }
                else if (extension == ".pdf")
                {
                    pageCount = this.ConvertPDF2Image(tempFilePath2, tempImageDirPath, string.Empty, 0, 0, imgFormat, imgResolution, imgQuality);
                }

                convertResult = ConvertResult.ConvertSuccess;
                return tempRelativeImageDirPath + "|" + pageCount;
            }
            catch (Exception ex)
            {
                // we need delete the image file when happen exception, because it maybe not correct
                try
                {
                    if (!string.IsNullOrEmpty(tempImageDirPath) && Directory.Exists(tempImageDirPath))
                    {
                        int count = 0;
                        while (true)
                        {
                            try
                            {
                                Directory.Delete(tempImageDirPath, true);
                            }
                            catch (Exception)
                            {
                                Thread.Sleep(50);
                            }

                            if (!Directory.Exists(tempImageDirPath) || count == 10)
                            {
                                break;
                            }
                            count++;
                        }
                    }
                }
                catch (Exception) { }

                throw ex;
            }
            finally
            {
                // we need delete the temp file when convert finished , if has exception ignore it.
                try
                {
                    // write operate log
                    WriteLog(logFilePath, "公文号：" + docId + "，开始时间：" + ConvertStartTime.ToLongTimeString() + "，结束时间：" +
DateTime.Now.ToLongTimeString() + "，耗时：" + (DateTime.Now - ConvertStartTime).Milliseconds + "毫秒" + "，转换结果：" + convertResult.ToString());

                    // delete temp file
                    if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                    if (!string.IsNullOrEmpty(tempFilePath2) && File.Exists(tempFilePath2))
                    {
                        File.Delete(tempFilePath2);
                    }
                }
                catch (Exception) { }
            }
        }
        /// <summary>
        /// PDF,Wrod转换成图片
        /// </summary>
        /// <param name="fileUrl">源文件路径</param>
        /// <param name="imgFormat">图片格式</param>
        /// <param name="imgResolution">图片分辨率</param>
        /// <param name="imgQuality">图片质量</param>
        /// <returns>返回值格式：图片存放目录|页数</returns>
        public string ConvertToImage(string fileUrl, ImageFormat imgFormat, int imgResolution, int imgQuality)
        {
            DateTime ConvertStartTime = DateTime.Now;
            ConvertResult convertResult = ConvertResult.HasException;
            int pageCount = 0;
            string docId = string.Empty;
            string tempFilePath = string.Empty;
            string tempFilePath2 = string.Empty;
            string tempFilePath3 = string.Empty;
            string tempImageDirPath = string.Empty;
            string extension = Path.GetExtension(fileUrl).ToLower();
            string logFilePath = Path.Combine(this._tempPath, "Log_" + ConvertStartTime.ToString("yyyyMMdd")) + ".txt";

            try
            {
                // validate parameter
                ValidateHelper.Begin().NotNullAndEmpty(fileUrl).NotNull(imgFormat)
                    .InRange(imgResolution, 1, 512).InRange(imgQuality, 1, 100).CheckFileType(extension);

                docId = Path.GetFileNameWithoutExtension(fileUrl);
                tempFilePath = Path.Combine(this._tempPath, Guid.NewGuid().ToString()) + extension;
                tempFilePath2 = Path.Combine(this._tempPath, Guid.NewGuid().ToString()) + extension;

                // if _autoDelOldFile is ture , we need clearup old file by date and ignore exception
                if (this._autoDelOldFile)
                {
                    ClearupDirDelegate clearupDirDelegate = ClearupDir;
                    clearupDirDelegate(this._tempPath, this._saveDirMaxFileCount, this._deleteDateTime);
                }

                // three case for copy file to local temp directory , 1.from Http:// , 2.from other PC , 3.from local
                if (fileUrl.StartsWith("http://") || fileUrl.StartsWith("https://"))
                {
                    DownloadToFile(fileUrl, tempFilePath);
                }
                else if (fileUrl.StartsWith(@"\\"))
                {
                    using (WindowsIdentityScope wis = new WindowsIdentityScope(this._serverIP, this._serverLogOnName, this._serverLogOnPassword))
                    {
                        File.Copy(fileUrl, tempFilePath);
                    }
                }
                else
                {
                    File.Copy(fileUrl, tempFilePath);
                }

                // if the file is ziped ,we need decompress it
                if (ZipFile.IsFileZiped(tempFilePath))
                {
                    ZipFile.DecompressFirstFile(tempFilePath, tempFilePath2);
                }
                else
                {
                    tempFilePath2 = tempFilePath;
                }

                string tempImageDirName = docId + "&" + File.GetLastWriteTime(tempFilePath2).ToFileTime().ToString();
                string tempRelativeImageDirPath = PathIncludeSlash(Path.Combine(this._relativeTempPath, tempImageDirName));
                tempImageDirPath = PathIncludeSlash(Path.Combine(this._tempPath, tempImageDirName));

                if (Directory.Exists(tempImageDirPath))
                {
                    DirectoryInfo tempImageDirInfo = new DirectoryInfo(tempImageDirPath);
                    if (tempImageDirInfo.GetFiles().Length > 0)
                    {
                        convertResult = ConvertResult.ReturnCashe;
                        return tempRelativeImageDirPath + "|" + tempImageDirInfo.GetFiles().Length;
                    }
                }

                // convert file to image
                if (extension == ".docx" || extension == ".doc")
                {
                    string tempPdfFileName = Guid.NewGuid().ToString() + ".pdf";
                    if (this.ConvertWord2PDF(tempFilePath2, this._tempPath, tempPdfFileName))
                    {
                        tempFilePath3 = Path.Combine(this._tempPath, tempPdfFileName);
                        pageCount = this.ConvertPDF2Image(tempFilePath3, tempImageDirPath, string.Empty, 0, 0, imgFormat, imgResolution, imgQuality);
                    }
                    else
                    {
                        pageCount = 0;
                    }
                }
                else if (extension == ".pdf")
                {
                    pageCount = this.ConvertPDF2Image(tempFilePath2, tempImageDirPath, string.Empty, 0, 0, imgFormat, imgResolution, imgQuality);
                }

                convertResult = ConvertResult.ConvertSuccess;
                return tempRelativeImageDirPath + "|" + pageCount;
            }
            catch (Exception ex)
            {
                // we need delete the image file when happen exception, because it maybe not correct
                try
                {
                    if (!string.IsNullOrEmpty(tempImageDirPath) && Directory.Exists(tempImageDirPath))
                    {
                        int count = 0;
                        while (true)
                        {
                            try
                            {
                                Directory.Delete(tempImageDirPath, true);
                            }
                            catch (Exception)
                            {
                                Thread.Sleep(50);
                            }

                            if (!Directory.Exists(tempImageDirPath) || count == 10)
                            {
                                break;
                            }
                            count++;
                        }
                    }
                }
                catch (Exception) { }

                throw ex;
            }
            finally
            {
                // we need delete the temp file when convert finished , if has exception ignore it.
                try
                {
                    // write operate log
                    WriteLog(logFilePath, "公文号：" + docId + "，开始时间：" + ConvertStartTime.ToLongTimeString() + "，结束时间：" +
DateTime.Now.ToLongTimeString() + "，耗时：" + (DateTime.Now - ConvertStartTime).Milliseconds + "毫秒" + "，转换结果：" + convertResult.ToString());

                    // delete temp file
                    if (!string.IsNullOrEmpty(tempFilePath) && File.Exists(tempFilePath))
                    {
                        File.Delete(tempFilePath);
                    }
                    if (!string.IsNullOrEmpty(tempFilePath2) && File.Exists(tempFilePath2))
                    {
                        File.Delete(tempFilePath2);
                    }
                    if (!string.IsNullOrEmpty(tempFilePath3) && File.Exists(tempFilePath3))
                    {
                        File.Delete(tempFilePath3);
                    }
                }
                catch (Exception) { }
            }
        }

        // convert to image 
        /// <summary>
        /// 
        /// </summary>
        /// <param name="srcFilePath"></param>
        /// <param name="destFileDir"></param>
        /// <param name="destFileName"></param>
        /// <param name="startPageNum"></param>
        /// <param name="endPageNum"></param>
        /// <param name="imgFormat"></param>
        /// <param name="imgResolution"></param>
        /// <param name="imgQuality"></param>
        /// <returns></returns>
        private int ConvertPDF2Image(string srcFilePath, string destFileDir, string destFileName,
            int startPageNum, int endPageNum, ImageFormat imgFormat, int imgResolution, int imgQuality)
        {
            // validate path
            ValidateHelper.Begin().NotNullAndEmpty(srcFilePath).FileExist(srcFilePath).NotNullAndEmpty(destFileDir)
                .NotNull(imgFormat).InRange(imgResolution, 1, 512).InRange(imgQuality, 1, 100);

            string imageExtention = "." + imgFormat.ToString();
            if (imgFormat == ImageFormat.Jpeg) { imageExtention = ".jpg"; }
            if (!string.IsNullOrEmpty(destFileName)) { destFileName = destFileName + "_"; }
            if (!Directory.Exists(destFileDir)) { Directory.CreateDirectory(destFileDir); }

            // open pdf file
            using (PDFFile pdfFile = PDFFile.Open(srcFilePath))
            {
                if (startPageNum <= 0) { startPageNum = 1; }
                if (endPageNum > pdfFile.PageCount || endPageNum <= 0) { endPageNum = pdfFile.PageCount; }
                if (startPageNum > endPageNum) { int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = tempPageNum; }

                // init encoder parameter
                EncoderParameter encoderParam = new EncoderParameter(Encoder.Quality, (long)imgQuality);
                EncoderParameters encoderParams = new EncoderParameters(1);
                encoderParams.Param[0] = encoderParam;
                ImageCodecInfo codecinfo = GetEncoderInfo("image/jpeg");

                // start to convert each page
                try
                {
                    for (int i = startPageNum - 1; i < endPageNum; i++)
                    {
                        using (Bitmap pdfBitmap = pdfFile.GetPageImage(i, (float)imgResolution))
                        {
                            pdfBitmap.Save(Path.Combine(destFileDir, destFileName + i.ToString() + imageExtention), codecinfo, encoderParams);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }

                return endPageNum - startPageNum + 1;
            }
        }
        /// <summary>
        /// WORD转换成图片
        /// </summary>
        /// <param name="fileUrl">源文件路径</param>
        /// <param name="resolution">图片分辨率</param>
        /// <param name="imgQuality">图片质量</param>
        /// <returns>返回值格式：图片存放目录|页数</returns>
        private int ConvertWord2Image(string srcFilePath, string destFileDir, string destFileName,
            int startPageNum, int endPageNum, ImageFormat imgFormat, int imgResolution, int imgQuality)
        {
            // validate path
            ValidateHelper.Begin().NotNullAndEmpty(srcFilePath).FileExist(srcFilePath).NotNullAndEmpty(destFileDir)
                .NotNull(imgFormat).InRange(imgResolution, 1, 512).InRange(imgQuality, 1, 100);

            string imageExtention = "." + imgFormat.ToString();
            if (imgFormat == ImageFormat.Jpeg) { imageExtention = ".jpg"; }
            if (!string.IsNullOrEmpty(destFileName)) { destFileName = destFileName + "_"; }
            if (!Directory.Exists(destFileDir)) { Directory.CreateDirectory(destFileDir); }

            // open word file
            Aspose.Words.Document wordFile = new Aspose.Words.Document(srcFilePath);

            if (wordFile.HasRevisions) { wordFile.AcceptAllRevisions(); }
            if (startPageNum <= 0) { startPageNum = 1; }
            if (endPageNum > wordFile.PageCount || endPageNum <= 0) { endPageNum = wordFile.PageCount; }
            if (startPageNum > endPageNum) { int tempPageNum = startPageNum; startPageNum = endPageNum; endPageNum = tempPageNum; }

            ImageSaveOptions imageSaveOptions = new ImageSaveOptions(GetSaveFormat(imgFormat));
            imageSaveOptions.Resolution = (float)imgResolution;
            imageSaveOptions.JpegQuality = imgQuality;

            // start to convert each page
            try
            {
                for (int i = startPageNum - 1; i < endPageNum; i++)
                {
                    imageSaveOptions.PageIndex = i;
                    wordFile.Save(Path.Combine(destFileDir, destFileName + i.ToString() + imageExtention), imageSaveOptions);
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

            return endPageNum - startPageNum + 1;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="srcFilePath"></param>
        /// <param name="destFileDir"></param>
        /// <param name="pdfName"></param>
        /// <returns></returns>
        private bool ConvertWord2PDF(string srcFilePath, string destFileDir, string destFileName)
        {
            bool result;

            // validate path
            ValidateHelper.Begin().NotNullAndEmpty(srcFilePath).FileExist(srcFilePath).NotNullAndEmpty(destFileDir).NotNullAndEmpty(destFileName);

            if (!Directory.Exists(destFileDir)) { Directory.CreateDirectory(destFileDir); }
            if (Path.GetExtension(destFileName).ToLower() != ".pdf") { destFileName = destFileName + ".pdf"; }

            try
            {
                object paramSourceDocPath = srcFilePath;
                object paramMissing = Type.Missing;
                string paramExportFilePath = Path.Combine(destFileDir, destFileName);

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

                    result = true;
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

            return result;
        }

        // common method
        private static string PathIncludeSlash(string path)
        {
            if (!path.EndsWith(@"\") && !path.EndsWith("/"))
            {
                return (path + @"\");
            }
            return path;
        }
        private static void WriteLog(string logFilePath, string content)
        {
            try
            {
                ValidateHelper.Begin().NotNullAndEmpty(logFilePath).NotNullAndEmpty(content);

                FileMode fm;

                if (File.Exists(logFilePath))
                {
                    fm = FileMode.Append;
                }
                else
                {
                    fm = FileMode.Create;
                }

                using (FileStream fs = new FileStream(logFilePath, fm))
                {
                    using (StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.Default))
                    {
                        sw.WriteLine(content);
                    }
                }

            }
            catch (Exception)
            {
                throw new IOException("写入日志失败！");
            }
        }
        private static void DownloadToFile(string url, string saveName)
        {
            try
            {
                using (WebClient webClient = new WebClient())
                {
                    webClient.DownloadFile(url.Replace("\\", "/"), saveName);
                    //using (FileStream outputStream = new FileStream(saveName, FileMode.Create, FileAccess.Write))
                    //{
                    //    byte[] buffer = webClient.DownloadData(url.Replace("\\", "/"));
                    //    outputStream.Write(buffer, 0, buffer.Length);
                    //    outputStream.Flush();
                    //}
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        private static void ClearupDir(string directory, int dirMaxFileCount, DateTime deleteDateTime)
        {
            try
            {
                if (Directory.Exists(directory))
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(directory);
                    FileInfo[] files = dirInfo.GetFiles();
                    DirectoryInfo[] dirs = dirInfo.GetDirectories();

                    if ((files.Length + dirs.Length) > dirMaxFileCount)
                    {
                        var queryFile = from file in files orderby file.CreationTime select file;
                        foreach (var q in queryFile)
                        {
                            if (q.CreationTime < deleteDateTime)
                            {
                                try { File.Delete(q.FullName); }
                                catch (Exception) { }
                            }
                            else break;
                        }

                        var queryDir = from dir in dirs orderby dir.CreationTime select dir;
                        foreach (var q in queryDir)
                        {
                            if (q.CreationTime < deleteDateTime)
                            {
                                try { Directory.Delete(q.FullName, true); }
                                catch (Exception) { }
                            }
                            else break;
                        }
                    }
                }
            }
            catch (Exception) { }
        }
        private static SaveFormat GetSaveFormat(ImageFormat imageFormat)
        {
            switch (imageFormat.ToString().ToLower())
            {
                case "jpeg":
                    return SaveFormat.Jpeg;
                case "png":
                    return SaveFormat.Png;
                case "bmp":
                    return SaveFormat.Bmp;
                case "tiff":
                    return SaveFormat.Tiff;
                default:
                    return SaveFormat.Unknown;
            }
        }
        private static ImageCodecInfo GetEncoderInfo(string mimeType)
        {
            ImageCodecInfo[] encoders;
            encoders = ImageCodecInfo.GetImageEncoders();
            for (int i = 0; i < encoders.Length; ++i)
            {
                if (encoders[i].MimeType == mimeType)
                {
                    return encoders[i];
                }
            }
            return null;
        }

        // enum
        private enum ConvertResult { ConvertSuccess, HasException, ReturnCashe }
    }
}
