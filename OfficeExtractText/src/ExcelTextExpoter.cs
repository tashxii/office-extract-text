using Microsoft.Office.Interop.Excel;
using OfficeExtractTexts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace OfficeExtractText
{
    class ExcelTextExpoter
    {
        private string file;
        private Workbooks books;

        public ExcelTextExpoter(string file, Workbooks books)
        {
            this.file = file;
            this.books = books;
        }

        internal List<string> Export()
        {
            List<string> contents = new List<string>();
            using (var book = new ComWrapper<Excel.Workbook>(books.Open(file,
                    UpdateLinks: Excel.XlUpdateLinks.xlUpdateLinksNever,
                    ReadOnly: true,
                    IgnoreReadOnlyRecommended: true,
                    Editable: false)
                ))
            {
                List<string> sheetNames = new List<string>();
                List<string> tempFiles = new List<string>();
                bool success = false;
                try
                {
                    for (int i = 1; i <= book.ComObject.Worksheets.Count; i++)
                    {
                        using (var sheet = new ComWrapper<Excel._Worksheet>(book.ComObject.Worksheets[i]))
                        {
                            var sheetName = sheet.ComObject.Name;//Not after save, because sheet name will be changed after saving.
                            sheetNames.Add(sheetName);
                            var tempFile1 = Path.GetTempFileName();//for sheet
                            tempFiles.Add(tempFile1);
                            //Text in sheet.
                            sheetNames.Add(sheetName);
                            var tempFile2 = Path.GetTempFileName();//for shapes & comments
                            tempFiles.Add(tempFile2);
                            sheet.ComObject.SaveAs(tempFile1, FileFormat: Excel.XlFileFormat.xlCSV);
                            //Text in shapes & comments
                            List<string> otherContents = new List<string>();
                            foreach (Excel.Shape shape in sheet.ComObject.Shapes)
                            {
                                ExtractShapesContents(otherContents, shape);
                            }
                            foreach (Excel.Comment comment in sheet.ComObject.Comments)
                            {
                                otherContents.Add(comment.Author + ":" + comment.Text());
                            }
                            File.WriteAllLines(tempFile2, otherContents, Encoding.Default);
                            success = true;
                        }
                    }
                }
                finally
                {
                    book.ComObject.Close(false);
                    //merge contents after closing 
                    if (success)
                    {
                        int i = 0;
                        foreach (var tempFile in tempFiles)
                        {
                            if (i % 2 == 0)
                            {//sheet=n+0, shapes=n+1
                                contents.Add("[" + sheetNames[i] + "]");
                            }
                            i++;
                            var sheetContents = FileUtils.MergeTextContents(new string[] { tempFile });
                            File.Delete(tempFile);
                            contents.AddRange(sheetContents);
                        }
                    }
                    FileUtils.DeleteFiles(tempFiles.ToArray());
                }
                return contents;
            }
        }

        private void ExtractShapesContents(List<string> contents, Excel.Shape shape)
        {
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                //To check group or not, use only shape.Type or shape.AutoShapeType, 
                //because other ways like shape.GroupItem.Count & shape.Ungroup thow an exception when shape is not a group.
                var groupShapes = shape.GroupItems;
                foreach(Excel.Shape subShape in shape.GroupItems)
                {
                    ExtractShapesContents(contents, subShape);
                }
                //for (int i = 1; i <= groupShapes.Count; i++)
                //{
                //    ExtractShapesContents(contents, groupShapes.Item(i));
                //}
            }
            else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
            {
                foreach (Excel.Shape subShape in shape.CanvasItems)
                {
                    ExtractShapesContents(contents, subShape);
                }
            }
            else
            {
                var text = shape.TextFrame?.Characters()?.Text;
                if (!String.IsNullOrEmpty(text))
                {
                    contents.Add(text);
                }
            }
        }
    }
}
