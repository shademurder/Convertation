using System;
using System.Collections.Generic;
using System.IO;

namespace CaseAgile.OfficePublisher
{
    class Converter
    {
        private static readonly string formatSuffix = ".html";

        public static bool CanPublish(string inputFile)
        {
            try
            {
                if (string.IsNullOrEmpty(inputFile)) return false;
                if (inputFile.EndsWith(Format.Docx.Value) ||
                  inputFile.EndsWith(Format.Doc.Value) ||
                  inputFile.EndsWith(Format.Pdf.Value) ||
                  inputFile.EndsWith(Format.Ppt.Value) ||
                  inputFile.EndsWith(Format.Pptx.Value) ||
                  inputFile.EndsWith(Format.Xls.Value) ||
                  inputFile.EndsWith(Format.Xlsx.Value)) return true;
                return false;
            }
            catch { return false; }
        }

        public static bool PublishOfficeHTML(string inputFile, string outputFolder, Dictionary<string, string> parameters, out string log)
        {
            StringWriter stringWriter = new StringWriter();
            try
            {
                if (!File.Exists(inputFile))
                {
                    stringWriter.Write("Input file does not esists");
                    log = stringWriter.ToString();
                    return false;
                }
                if (!Directory.Exists(outputFolder))
                {
                    try
                    {
                        Directory.CreateDirectory(outputFolder);
                        stringWriter.WriteLine("Output folder created");
                    }
                    catch
                    {
                        stringWriter.Write("Error of creation output folder");
                        log = stringWriter.ToString();
                        return false;
                    }
                }
                if (inputFile.EndsWith(Format.Docx.Value) || inputFile.EndsWith(Format.Doc.Value))
                {
                    var fileInfo = new FileInfo(inputFile);
                    Document.ConvertToHtml(inputFile, outputFolder, fileInfo.Name + formatSuffix, parameters);
                    stringWriter.WriteLine("File convertation completed");
                    log = stringWriter.ToString();
                    return true;
                }
                else if (inputFile.EndsWith(Format.Pdf.Value))
                {
                    var fileInfo = new FileInfo(inputFile);
                    Pdf.ConvertToHtml(inputFile, outputFolder, fileInfo.Name + formatSuffix, parameters);
                    stringWriter.WriteLine("File convertation completed");
                    log = stringWriter.ToString();
                    return true;
                }
                else if (inputFile.EndsWith(Format.Ppt.Value) || inputFile.EndsWith(Format.Pptx.Value))
                {
                    var fileInfo = new FileInfo(inputFile);
                    Presentation.ConvertToHtml(inputFile, outputFolder, fileInfo.Name + formatSuffix, parameters);
                    stringWriter.WriteLine("File convertation completed");
                    log = stringWriter.ToString();
                    return true;
                }
                else if (inputFile.EndsWith(Format.Xls.Value) || inputFile.EndsWith(Format.Xlsx.Value))
                {
                    var fileInfo = new FileInfo(inputFile);
                    Excel.ConvertToHtml(inputFile, outputFolder, parameters);
                    stringWriter.WriteLine("File convertation completed");
                    log = stringWriter.ToString();
                    return true;
                }
                else
                {
                    stringWriter.WriteLine("Unknown format error");
                    log = stringWriter.ToString();
                    return false;
                }
            }
            catch(Exception e)
            {
                stringWriter.WriteLine("Convertation error");
                stringWriter.WriteLine(e.StackTrace);
                log = stringWriter.ToString();
                return false;
            }
        }        
    }
}
