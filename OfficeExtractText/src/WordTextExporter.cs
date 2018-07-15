using Microsoft.Office.Interop.Word;
using OfficeExtractTexts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace OfficeExtractText
{
    internal class WordTextExporter
    {
        private string file;
        private Documents docs;

        public WordTextExporter(string file, Documents docs)
        {
            this.file = file;
            this.docs = docs;
        }

        internal List<string> Export()
        {
            var contents = new List<string>();
            using (var doc = new ComWrapper<Word.Document>(
                docs.Open(file,
                    ReadOnly: true,
                    AddToRecentFiles: false,
                    Visible: false)
                ))
            {
                var tempFiles = new string[2];
                bool success = false;
                try
                {
                    tempFiles[0] = Path.GetTempFileName();
                    tempFiles[1] = Path.GetTempFileName();

                    //Text in word.
                    doc.ComObject.SaveAs2(tempFiles[0], FileFormat: Word.WdSaveFormat.wdFormatText);
                    //Text in shapes.
                    List<string> otherContents = new List<string>();
                    foreach (Word.Shape shape in doc.ComObject.Shapes)
                    {
                        ExtractShapeContents(otherContents, shape);
                    }
                    foreach (Word.Comment comment in doc.ComObject.Comments)
                    {
                        otherContents.Add(comment.Author + ":" + comment.Range.Text);
                    }
                    File.WriteAllLines(tempFiles[1], otherContents, Encoding.Default);
                    success = true;
                }
                finally
                {
                    doc.ComObject.Close(false);
                    //merge contents after closing doc.
                    if (success)
                    {
                        contents = FileUtils.MergeTextContents(tempFiles);
                    }
                    FileUtils.DeleteFiles(tempFiles);
                }
                return contents;
            }
        }
        private void ExtractShapeContents(List<string> contents, Word.Shape shape)
        {
            shape.Select();//shape.Type fails if not selected. This problem is word only.
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                //To check group or not, use only shape.AutoShapeType == msoShapeMixed or shape.Type == msoGroup,
                //because other ways like shape.GroupItem.Count & shape.Ungroup thow an exception when shape is not a group.
                foreach (Word.Shape subShape in shape.GroupItems)
                {
                    ExtractShapeContents(contents, subShape);
                }
            }
            else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
            {
                foreach (Word.Shape subShape in shape.CanvasItems)
                {
                    ExtractShapeContents(contents, subShape);
                }
            }
            else
            {
                if (shape.TextFrame != null && shape.TextFrame.HasText != 0)
                {
                    var text = shape.TextFrame?.TextRange?.Text;
                    if (!String.IsNullOrEmpty(text))
                    {
                        contents.Add(text);
                    }
                }
            }
        }
    }
}