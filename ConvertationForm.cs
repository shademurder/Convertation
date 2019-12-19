using CaseAgile.OfficePublisher;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Text;

namespace Convertation
{
    public partial class ConvertationForm : Form
    {
        public ConvertationForm()
        {
            InitializeComponent();
            Resize += OnResize;
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Stopwatch stopwatch = new System.Diagnostics.Stopwatch();
            LogBox.Text = "";
            try
            {
                stopwatch.Start();
                if (!Directory.Exists(InputDirectory.Text))
                {
                    MessageBox.Show("Correctly fill the input directory");
                }
                else
                {
                    string log = ConvertDirectoryToHtml(InputDirectory.Text, OutputDirectory.Text);
                    LogBox.Text = log;
                }
                stopwatch.Stop();
                LogBox.Text += Environment.NewLine + stopwatch.Elapsed.ToString();
            }
            catch (Exception exception)
            {
                LogBox.Text = exception.StackTrace;
            }
        }

        private void OnResize(object sender, EventArgs args)
        {
            LogBox.Size = new Size(Size.Width - 40, Size.Height - 125);
            Process.Location = new Point(Size.Width - 103, Process.Location.Y);
            InputDirectory.Size = new Size(Size.Width - 221, InputDirectory.Size.Height);
            OutputDirectory.Size = new Size(Size.Width - 221, OutputDirectory.Size.Height);
        }

        private void SetInput_Click(object sender, EventArgs e)
        {
            FolderDialog.ShowDialog();
            InputDirectory.Text = FolderDialog.SelectedPath;
        }

        private void SetOutput_Click(object sender, EventArgs e)
        {
            FolderDialog.ShowDialog();
            OutputDirectory.Text = FolderDialog.SelectedPath; 
        }

        private void ConvertationForm_Load(object sender, EventArgs e)
        {
            //EvoPdf.PdfToHtml.PdfToHtmlConverter pdfToHtmlConverter = new EvoPdf.PdfToHtml.PdfToHtmlConverter();
            //pdfToHtmlConverter.ConvertPdfPagesToHtmlFile("F:\\TestFiles\\11.1.11\\11.1.8.2_Fibre_Splicing_Guidelines.pdf", "F:\\TestFiles2\\11.1.11\\2\\", "11.1.8.2_Fibre_Splicing_Guidelines.pdf.html");
            //Winnovative.PdfToHtml.PdfToHtmlConverter pdfToHtmlConverter = new Winnovative.PdfToHtml.PdfToHtmlConverter();
            //pdfToHtmlConverter.ConvertPdfPagesToHtmlFile("F:\\TestFiles\\11.1.11\\11.1.8.2_Fibre_Splicing_Guidelines.pdf", "F:\\TestFiles2\\11.1.11\\2\\", "11.1.8.2_Fibre_Splicing_Guidelines.pdf.html");
            //SautinSoft.PdfFocus pdfToHtmlConverter = new SautinSoft.PdfFocus();
            //pdfToHtmlConverter.OpenPdf("F:\\TestFiles\\11.1.11\\11.1.8.2_Fibre_Splicing_Guidelines.pdf");
            //pdfToHtmlConverter.HtmlOptions.IncludeImageInHtml = true;
            //string html = pdfToHtmlConverter.ToHtml();
            //Bytescout.PDF2HTML.HTMLExtractor extractor = new Bytescout.PDF2HTML.HTMLExtractor();
            //extractor.LoadDocumentFromFile("F:\\TestFiles\\11.1.11\\11.1.8.2_Fibre_Splicing_Guidelines.pdf");
            //extractor.SaveHtmlToFile("F:\\TestFiles2\\11.1.11\\3\\11.1.8.2_Fibre_Splicing_Guidelines.pdf.html");

            //extractor.Dispose();
            //Spire.Pdf.PdfDocument pdf = new Spire.Pdf.PdfDocument("F:\\TestFiles\\11.1.11\\11.1.8.2_Fibre_Splicing_Guidelines.pdf");

            //"F:\\TestFiles2\\11.1.11\\4\\11.1.8.2_Fibre_Splicing_Guidelines.pdf.docx"
            //pdf.SaveToFile("F:\\TestFiles2\\11.1.11\\5\\11.1.8.2_Fibre_Splicing_Guidelines.pdf.docx", Spire.Pdf.FileFormat.);
            //string result = System.Text.Encoding.UTF8.GetString(stream.ToArray());
            //File.WriteAllText("F:\\TestFiles2\\11.1.11\\5\\11.1.8.2_Fibre_Splicing_Guidelines.pdf.docx", result, Encoding.UTF8);


            //string log = ConvertDirectoryToHtml("F:\\TestFiles\\7112019", "F:\\TestFiles2\\7112019");

            //LogBox.Text = log;

            //Spire.Doc.Document doc = new Spire.Doc.Document("F:\\TestFiles\\221119\\3.5.1.19_Figure_8_Technical_Specifications_Sheets.docx");
            //doc.SaveToFile("F:\\TestFiles2\\221119\\3.5.1.19_Figure_8_Technical_Specifications_Sheets.docx.html", Spire.Doc.FileFormat.Html);
            //int a = 0;
        }

        private string ConvertDirectoryToHtml(string inputDir, string outputDir)
        {
            string log = "";
            try
            {
                if (!Directory.Exists(inputDir))
                {
                    throw new Exception("Input directory does not esists");
                }
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                Dictionary<string, string> parameters = new Dictionary<string, string>
                {
                    { Parameter.ExcelMaxRows.Value, "100" },
                    { Parameter.ExcelMaxColumns.Value, "100" },
                    { Parameter.ExcelMaxSheetSize.Value, "100" }
                };
                DirectoryInfo directory = new DirectoryInfo(inputDir);
                foreach (FileInfo fileInfo in directory.GetFiles())
                {
                    if (fileInfo.Name.EndsWith(Format.Pdf.Value) || fileInfo.Name.EndsWith(Format.Doc.Value) || fileInfo.Name.EndsWith(Format.Docx.Value) || fileInfo.Name.EndsWith(Format.Ppt.Value) || fileInfo.Name.EndsWith(Format.Pptx.Value) || fileInfo.Name.EndsWith(Format.Xls.Value) || fileInfo.Name.EndsWith(Format.Xlsx.Value))
                    {
                        bool result = Converter.PublishOfficeHTML(fileInfo.FullName, outputDir + "/" + fileInfo.Name, parameters, out log);
                    }
                }
                foreach (DirectoryInfo directoryInfo in directory.GetDirectories())
                {
                    ConvertDirectoryToHtml(directoryInfo.FullName, outputDir + "/" + directoryInfo.Name);
                }
            }
            catch (Exception e)
            {
                log += Environment.NewLine + "Error: " + e.Message;
                MessageBox.Show("Error: " + e.StackTrace);
            }
            return log;
        }
    }
}
