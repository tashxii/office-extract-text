using OfficeExtractTexts;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace OfficeExtractText
{
    class PowerPointTextExporter
    {
        private string file;
        private PowerPoint.Presentations ppts;

        public PowerPointTextExporter(string file, PowerPoint.Presentations ppts)
        {
            this.file = file;
            this.ppts = ppts;
        }

        internal List<string> Export()
        {
            var contents = new List<string>();
            using (var ppt = new ComWrapper<PowerPoint.Presentation>(
                ppts.Open(file,
                    ReadOnly: Microsoft.Office.Core.MsoTriState.msoTrue,
                    WithWindow: Microsoft.Office.Core.MsoTriState.msoFalse)
                ))
            {
                var tempFiles = new string[2];
                var success = false;
                try
                {
                    tempFiles[0] = Path.GetTempFileName();
                    tempFiles[1] = Path.GetTempFileName();

                    //Text in PPT
                    //Save as rich text file.
                    ppt.ComObject.SaveAs(tempFiles[0], FileFormat: PowerPoint.PpSaveAsFileType.ppSaveAsRTF);
                    //Read and save as a text file.
                    string richText = File.ReadAllText(tempFiles[0], Encoding.Default);
                    //Cheep trick to convert text from rtf.
                    RichTextBox richTextBox = new RichTextBox();
                    richTextBox.Rtf = richText;
                    File.WriteAllText(tempFiles[0], richTextBox.Text, Encoding.Default);

                    //Text in shapes & comments
                    var slideContents = new List<string>();
                    foreach (PowerPoint.Slide slide in ppt.ComObject.Slides)
                    {
                        foreach (PowerPoint.Shape shape in slide.Shapes)
                        {
                            ExtractShapeContents(slideContents, shape);
                        }
                        foreach (PowerPoint.Comment comment in slide.Comments)
                        {
                            slideContents.Add(comment.Author + ":" + comment.Text);
                        }
                        slideContents.Add(slide.NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text);//placefolders[1] is slide itself.
                    }
                    File.WriteAllLines(tempFiles[1], slideContents, Encoding.Default);
                    success = true;
                }
                finally
                {
                    ppt.ComObject.Close();
                    //merge contents after closing ppt.
                    if(success)
                    {
                        contents = FileUtils.MergeTextContents(tempFiles);
                    }
                    FileUtils.DeleteFiles(tempFiles);
                }
                return contents;
            }
        }

        private void ExtractShapeContents(List<string> contents, PowerPoint.Shape shape)
        {
            if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                //To check group or not, use only shape.AutoShapeType == msoShapeMixed or shape.Type == msoGroup,
                //because other ways like shape.GroupItem.Count & shape.Ungroup thow an exception when shape is not a group.
                foreach (PowerPoint.Shape subShape in shape.GroupItems)
                {
                    ExtractShapeContents(contents, subShape);
                }
            }
            else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
            {
                foreach (PowerPoint.Shape subShape in shape.CanvasItems)
                {
                    ExtractShapeContents(contents, subShape);
                }
            }
            else
            {
                if (shape.TextFrame != null && shape.TextFrame.HasText == Microsoft.Office.Core.MsoTriState.msoTrue)
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
