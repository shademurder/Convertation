using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using System;
using System.IO;
using System.Text;
using System.Collections.Generic;

namespace CaseAgile.OfficePublisher
{
    class Document
    {
        private static readonly string imagePrefix = "src=\"data:image/png;base64,";
        private static readonly string replaceValue = "Evaluation Warning: The document was created with Spire.Doc for .NET.</span></p>";
        private static readonly string replaceValueDoc = "<p class=\"Normal\"><span style=\"color:#FF0000;font-size:12pt;font-family:Calibri;mso-fareast-font-family:Calibri;mso-bidi-font-family:'Times New Roman';\">Evaluation Warning: The document was created with Spire.Doc for .NET.</span></p>";
        private static readonly string replaceValueDocx = "<p class=\"Normal\"><span style=\"color:#FF0000;font-size:12pt;font-family:Calibri;mso-fareast-font-family:宋体;mso-bidi-font-family:Arial;lang:RU-RU;mso-fareast-language:EN-US;mso-ansi-language:AR-SA;\">Evaluation Warning: The document was created with Spire.Doc for .NET.</span></p>";
        /// <summary>
        /// Convert doc or docx to html with base64 pictures
        /// </summary>
        /// <param name="inputFileName">Absolute name of input file</param>
        /// <param name="outputFileName">Absolute name of output file</param>
        public static void ConvertToHtml(string inputFileName, string outputFileName)
        {
            Spire.Doc.Document document = new Spire.Doc.Document(inputFileName);
            using (var stream = new MemoryStream())
            {
                string result = System.Text.Encoding.UTF8.GetString(stream.ToArray());
                document.SaveToStream(stream, Spire.Doc.FileFormat.Html);
                if (result.IndexOf(replaceValueDocx) < 0 && result.IndexOf(replaceValueDoc) < 0)
                {
                    result = System.Text.Encoding.UTF8.GetString(stream.ToArray()).Replace(replaceValueDocx, "").Replace(replaceValueDoc, "").Replace(replaceValue, "");
                }
                else
                {
                    result = System.Text.Encoding.UTF8.GetString(stream.ToArray()).Replace(replaceValueDocx, "").Replace(replaceValueDoc, "");
                }
                result = ReplaceImages(document, result);
            
                File.WriteAllText(outputFileName, result, Encoding.UTF8);
                result = null;
            }
            document.Dispose();
        }

        private static string ReplaceImages(Spire.Doc.Document document, string html)
        {
            int index = 1;
            foreach (Section section in document.Sections)
            {
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    foreach (DocumentObject docObject in paragraph.ChildObjects)
                    {
                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture picture = docObject as DocPicture;
                            using (var memoryStream = new MemoryStream())
                            {
                                picture.Image.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                                string base64Image = System.Convert.ToBase64String(memoryStream.ToArray());
                                html = html.Replace("src=\"_images/_img" + index + ".png\"", imagePrefix + base64Image + "\"").
                                            Replace("src=\"_images/_img" + index + ".jpeg\"", imagePrefix + base64Image + "\"");
                                index++;
                            }
                        }
                    }
                }
            }
            return html;
        }

        /// <summary>
        /// Convert doc or docx to html with picture files
        /// </summary>
        /// <param name="inputFileName">Absolute name of input file</param>
        /// <param name="outputFolder">Folder for saving results</param>
        /// <param name="newFileName">Relative filename of new html</param>
        public static void ConvertToHtml(string inputFileName, string outputFolder, string newFileName, Dictionary<string, string> parameters)
        {
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }
            Spire.Doc.Document document = new Spire.Doc.Document(inputFileName);
            using (var stream = new MemoryStream())
            {
                
                document.SaveToStream(stream, Spire.Doc.FileFormat.Html);
                string result = System.Text.Encoding.UTF8.GetString(stream.ToArray());
                if (result.IndexOf(replaceValueDocx) < 0 && result.IndexOf(replaceValueDoc) < 0)
                {
                    result = System.Text.Encoding.UTF8.GetString(stream.ToArray()).Replace(replaceValueDocx, "").Replace(replaceValueDoc, "").Replace(replaceValue, "");
                }
                else
                {
                    result = System.Text.Encoding.UTF8.GetString(stream.ToArray()).Replace(replaceValueDocx, "").Replace(replaceValueDoc, "");
                }
                AddImages(document, outputFolder, ref result);
                File.WriteAllText(outputFolder + "/" + newFileName, result, Encoding.UTF8);
                result = null;
            }
            document.Dispose();
        }

        public static void AddImages(Spire.Doc.Document document, string outputFolder, ref string result)
        {
            if(!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }
            List<DocPicture> attachments = new List<DocPicture>();
            List<string> attachmentNames = new List<string>();
            int index = 1;
            foreach (Section section in document.Sections)
            {
                //OLE
                foreach (DocumentObject documentObject in section.Body.ChildObjects)
                {
                    if (documentObject is Paragraph)
                    {
                        Paragraph par = documentObject as Paragraph;
                        foreach (DocumentObject docObj in par.ChildObjects)
                        {
                            if (docObj.DocumentObjectType == DocumentObjectType.OleObject)
                            {
                                DocOleObject Ole = docObj as DocOleObject;
                                //Excel.Sheet.8 is for XLS format
                                //Excel.Sheet.12 is for XLSX format
                                //AcroExch.Document.11 is for PDF
                                //AcroExch.Document.7 is for PDF
                                //Word.Document.8 is for Doc format
                                //Word.Document.12 is for Docx format
                                if (Ole.ObjectType == "AcroExch.Document.11" || Ole.ObjectType == "AcroExch.Document.7")
                                {
                                    attachments.Add(Ole.OlePicture);
                                    attachmentNames.Add("Ole" + attachments.Count + ".pdf");
                                    File.WriteAllBytes(outputFolder + "/Ole" + attachments.Count + ".pdf", Ole.NativeData);
                                }
                                else if (Ole.ObjectType == "Excel.Sheet.8")
                                {
                                    attachments.Add(Ole.OlePicture);
                                    attachmentNames.Add("Ole" + attachments.Count + ".xls");
                                    File.WriteAllBytes(outputFolder + "/Ole" + attachments.Count + ".xls", Ole.NativeData);
                                }
                                else if (Ole.ObjectType == "Excel.Sheet.12")
                                {
                                    attachments.Add(Ole.OlePicture);
                                    attachmentNames.Add("Ole" + attachments.Count + ".xlsx");
                                    File.WriteAllBytes(outputFolder + "/Ole" + attachments.Count + ".xlsx", Ole.NativeData);
                                }
                                else if(Ole.ObjectType == "Word.Document.8")
                                {
                                    attachments.Add(Ole.OlePicture);
                                    attachmentNames.Add("Ole" + attachments.Count + ".doc");
                                    File.WriteAllBytes(outputFolder + "/Ole" + attachments.Count + ".doc", Ole.NativeData);
                                }
                                else if (Ole.ObjectType == "Word.Document.12")
                                {
                                    attachments.Add(Ole.OlePicture);
                                    attachmentNames.Add("Ole" + attachments.Count + ".docx");
                                    File.WriteAllBytes(outputFolder + "/Ole" + attachments.Count + ".docx", Ole.NativeData);
                                }
                            }
                        }
                    }
                }
                //Header
                foreach (Paragraph paragraph in section.HeadersFooters.Header.Paragraphs)
                {
                    foreach (DocumentObject docObject in paragraph.ChildObjects)
                    {
                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture picture = docObject as DocPicture;
                            picture.Image.Save(outputFolder + "/img" + index + ".png", System.Drawing.Imaging.ImageFormat.Png);
                            result = result.Replace("src=\"_images/_img" + index + ".png\"", "src=\"img" + index + ".png\"").
                                            Replace("src=\"_images/_img" + index + ".jpeg\"", "src=\"img" + index + ".png\"");
                            index++;
                        }
                    }
                }
                //Body
                foreach (Paragraph paragraph in section.Paragraphs)
                {
                    if(paragraph.PreviousSibling is Table)
                    {
                        Table table = paragraph.PreviousSibling as Table;
                        for (int r = 0; r < table.Rows.Count; r++)
                        {
                            for (int c = 0; c < table.Rows[r].Cells.Count; c++)
                            {
                                foreach (Paragraph par in table.Rows[r].Cells[c].Paragraphs)
                                {
                                    foreach (DocumentObject docObject in par.ChildObjects)
                                    {
                                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                                        {
                                            DocPicture picture = docObject as DocPicture;
                                            picture.Image.Save(outputFolder + "/img" + index + ".png", System.Drawing.Imaging.ImageFormat.Png);
                                            result = result.Replace("src=\"_images/_img" + index + ".png\"", "src=\"img" + index + ".png\"").
                                                            Replace("src=\"_images/_img" + index + ".jpeg\"", "src=\"img" + index + ".png\"");
                                            if (attachments.Contains(picture))
                                            {
                                                try
                                                {
                                                    string href = GetAttachmentHref(attachments, attachmentNames, picture);
                                                    int imgStartIndex = result.Substring(0, result.IndexOf("src=\"img" + index + ".png\"")).LastIndexOf("<img");
                                                    int imgEndIndex = result.IndexOf(">", imgStartIndex);
                                                    string imageText = result.Substring(imgStartIndex, imgEndIndex - imgStartIndex + 1);
                                                    result = result.Replace(imageText, "<a href=\"" + href + "\">" + imageText + "</a>");
                                                }
                                                catch (Exception) { }
                                            }
                                            index++;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    foreach (DocumentObject docObject in paragraph.ChildObjects)
                    {
                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture picture = docObject as DocPicture;
                            picture.Image.Save(outputFolder + "/img" + index + ".png", System.Drawing.Imaging.ImageFormat.Png);
                            result = result.Replace("src=\"_images/_img" + index + ".png\"", "src=\"img" + index + ".png\"").
                                            Replace("src=\"_images/_img" + index + ".jpeg\"", "src=\"img" + index + ".png\"");
                            if(attachments.Contains(picture))
                            {
                                try
                                {
                                    string href = GetAttachmentHref(attachments, attachmentNames, picture);
                                    int imgStartIndex = result.Substring(0, result.IndexOf("src=\"img" + index + ".png\"")).LastIndexOf("<img");
                                    int imgEndIndex = result.IndexOf(">", imgStartIndex);
                                    string imageText = result.Substring(imgStartIndex, imgEndIndex - imgStartIndex + 1);
                                    result = result.Replace(imageText, "<a href=\"" + href + "\">" + imageText + "</a>");
                                }
                                catch (Exception) { }
                            }
                            index++;
                        }
                    }
                }
                //Footer
                foreach (Paragraph paragraph in section.HeadersFooters.Footer.Paragraphs)
                {
                    foreach (DocumentObject docObject in paragraph.ChildObjects)
                    {
                        if (docObject.DocumentObjectType == DocumentObjectType.Picture)
                        {
                            DocPicture picture = docObject as DocPicture;
                            picture.Image.Save(outputFolder + "/img" + index + ".png", System.Drawing.Imaging.ImageFormat.Png);
                            result = result.Replace("src=\"_images/_img" + index + ".png\"", "src=\"img" + index + ".png\"").
                                            Replace("src=\"_images/_img" + index + ".jpeg\"", "src=\"img" + index + ".png\"");
                            index++;
                        }
                    }
                }
            }
        }
        private static string GetAttachmentHref(List<DocPicture> attachments, List<string> attachmentNames, DocPicture picture)
        {
            try
            {
                for (int i = 0; i < attachments.Count; i++)
                {
                    if (attachments[i] == picture)
                    {
                        return attachmentNames[i];
                    }
                }
            }
            catch (Exception) { }
            return "";
        }
    }
}
