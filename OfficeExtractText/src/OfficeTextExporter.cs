using OfficeExtractText.Properties;
using OfficeExtractTexts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeExtractText
{
    class OfficeTextExporter
    {
        private static string[] EXCLE_EXTENSIONS = { ".xls", ".xlsx", ".xlsm" };
        private static string[] WORD_EXTENSIONS = { ".doc", ".docx", ".docm" };
        private static string[] POWERPOINT_EXTENSIONS = { ".ppt", ".pptx", ".pptm" };

        private string targetPath;
        private OfficeTextExporterOption options;
        private bool hasWarning = false;
        internal bool HasWarning() { return hasWarning; }

        public OfficeTextExporter(string targetPath, OfficeTextExporterOption options)
        {
            this.targetPath = targetPath;
            this.options = options;
        }

        internal void Execute()
        {
            //Check target files.
            var files = GetTargetFiles();
            if (files.Count == 0)
            {
                throw new ApplicationException(String.Format(Resource.NO_TARGET_FOUND));
            }
            //Init excel, word, powerpoint.
            using (var excel = new ComWrapper<Excel.Application>(new Excel.Application() { Visible = false, DisplayAlerts = false }))
            using (var books = new ComWrapper<Excel.Workbooks>(excel.ComObject.Workbooks))
            using (var word = new ComWrapper<Word.Application>(new Word.Application() { Visible = false, DisplayAlerts = Word.WdAlertLevel.wdAlertsNone }))
            using (var docs = new ComWrapper<Word.Documents>(word.ComObject.Documents))
            using (var powerPoint = new ComWrapper<PowerPoint.Application>(new PowerPoint.Application() { DisplayAlerts = PowerPoint.PpAlertLevel.ppAlertsNone }))
            using (var ppts = new ComWrapper<PowerPoint.Presentations>(powerPoint.ComObject.Presentations))
            {
                try
                {
                    foreach (var file in files)
                    {
                        try
                        {
                            ConsoleLogger.WriteLog(String.Format(Resource.EXTRACTING_TEXT_PATH_ARG0, file));
                            if (IsExcel(file))
                            {
                                ExportExcel(file, books);
                            }
                            else if (IsWord(file))
                            {
                                ExportWord(file, docs);
                            }
                            else if (IsPowerPoint(file))
                            {
                                ExportPowerPoint(file, ppts);
                            }
                        }
                        catch (COMException e)
                        {
                            ConsoleLogger.WriteWarning(String.Format(Resource.SKIPPED_EXTRACTING_TEXT_PATH_ARG0_DETAILS_ARG1, file, e.ToString()));
                            this.hasWarning = true;
                        }
                    }
                }
                finally
                {
                    try { books.ComObject.Close(); } catch (Exception e) { ConsoleLogger.WriteError(e.ToString()); }
                    try { excel.ComObject.Quit(); } catch (Exception e) { ConsoleLogger.WriteError(e.ToString()); }
                    //No need to close docs, because there is no opened file.(And an error occurrs when closing here.)
                    try { word.ComObject.Quit(); } catch (Exception e) { ConsoleLogger.WriteError(e.ToString()); }
                    try { powerPoint.ComObject.Quit(); } catch (Exception e) { ConsoleLogger.WriteError(e.ToString()); }
                }
            }
        }

        private List<string> GetTargetFiles()
        {
            var extensions = GetTargetExtensions();
            var result = new List<string>();
            if (File.Exists(targetPath))
            {
                if (!targetPath.ToLower().StartsWith("~$") && IsTargetFile(targetPath, extensions))
                {
                    result.Add(targetPath);
                    return result;
                }
                else
                {
                    throw new ApplicationException(String.Format(Resource.SPECIFIED_FILE_PATH_IS_NOT_TARGET_FILE_PATH_ARG0, targetPath));
                }
            }
            else if (Directory.Exists(targetPath))
            {
                SearchOption searchOption = (options.ExtractSubDir) ? SearchOption.AllDirectories : SearchOption.TopDirectoryOnly;
                var files = Directory.EnumerateFiles(targetPath, "*", searchOption)
                    .Where(s => !Path.GetFileName(s).ToLower().StartsWith("~$") && IsTargetFile(s, extensions));
                result.AddRange(files);
                if (result.Count == 0)
                {
                    throw new ApplicationException(String.Format(Resource.NO_TARGET_FILE_IN_SPECIFIED_DIRECTORY_PATH_ARG0, targetPath));
                }
                return result;
            }
            else
            {
                throw new ApplicationException(String.Format(Resource.SPECIFIED_FILE_PATH_IS_NOT_TARGET_FILE_PATH_ARG0, targetPath));
            }
        }

        private bool IsExcel(string file)
        {
            return EXCLE_EXTENSIONS.Contains(Path.GetExtension(file));
        }

        private bool IsWord(string file)
        {
            return WORD_EXTENSIONS.Contains(Path.GetExtension(file));
        }

        private bool IsPowerPoint(string file)
        {
            return POWERPOINT_EXTENSIONS.Contains(Path.GetExtension(file));
        }

        private List<string> GetTargetExtensions()
        {
            var extensions = new List<string>();
            if (options.ExtractExcel)
            {
                extensions.AddRange(EXCLE_EXTENSIONS);
            }
            if (options.ExtractWord)
            {
                extensions.AddRange(WORD_EXTENSIONS);
            }
            if (options.ExtractPowerPoint)
            {
                extensions.AddRange(POWERPOINT_EXTENSIONS);
            }
            return extensions;
        }

        private bool IsTargetFile(string file, List<string> extentions)
        {
            foreach (var extention in extentions)
            {
                if (file.ToLower().EndsWith(extention))
                {
                    return true;
                }
            }
            return false;
        }

        public string GetSavePath(string file)
        {
            string textPath = Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file) + ".txt");
            string baseDir = Path.GetDirectoryName(Path.GetFullPath(targetPath));
            textPath = options.OutputDir + textPath.Substring(baseDir.Length);
            Directory.CreateDirectory(Path.GetDirectoryName(textPath));
            return textPath;
        }

        private void ExportExcel(string file, ComWrapper<Excel.Workbooks> books)
        {
            var exporter = new ExcelTextExpoter(file, books.ComObject);
            List<string> contents = exporter.Export();
            if (contents.Count != 0)
            {
                File.WriteAllLines(GetSavePath(file), contents, Encoding.Default);
            }
        }

        private void ExportWord(string file, ComWrapper<Word.Documents> docs)
        {
            var exporter = new WordTextExporter(file, docs.ComObject);
            var contents = exporter.Export();
            if (contents.Count != 0)
            {
                File.WriteAllLines(GetSavePath(file), contents, Encoding.Default);
            }
        }

        private void ExportPowerPoint(string file, ComWrapper<PowerPoint.Presentations> ppts)
        {
            var exporter = new PowerPointTextExporter(file, ppts.ComObject);
            List<string> contents = exporter.Export();
            if (contents.Count != 0)
            {
                File.WriteAllLines(GetSavePath(file), contents, Encoding.Default);
            }
        }
    }
}