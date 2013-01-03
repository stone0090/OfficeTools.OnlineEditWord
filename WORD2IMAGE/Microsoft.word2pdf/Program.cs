using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Microsoft.Office.Interop.Word;

namespace Microsoft.word2pdf
{
    class Program
    {
        /// <summary>
        /// this function copy from Microsoft MSDN
        /// </summary>
        /// <param name="wordInputPath"></param>
        /// <param name="pdfOutputPath"></param>
        /// <param name="pdfName"></param>
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

        static void Main(string[] args)
        {
            ConvertWord2PDF("E:\\111.doc","E:","111");
        }
    }
}
