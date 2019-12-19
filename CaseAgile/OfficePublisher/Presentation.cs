using System.IO;
using System.Text;
using System.Collections.Generic;

namespace CaseAgile.OfficePublisher
{
    class Presentation
    {
        private static readonly string replaceValuePptx = "<text font-family=\"Arial\" font-weight=\"bold\" font-size=\"7.5pt\" fill=\"#ffd8cf\" opacity=\"0.75\"><tspan x=\"10\" y=\"20\" textLength=\"388.950134\" id=\"tp_0\">Evaluation Warning : The document was created with  Spire.Presentation for .NET</tspan></text>";
        private static readonly string replaceValue = "Evaluation Warning : The document was created with  Spire.Presentation for .NET";
        /// <summary>
        /// Convert presentation to html with base64 pictures
        /// </summary>
        /// <param name="inputFileName">Absolute name of input file</param>
        /// <param name="outputFileName">Absolute name of output file</param>
        public static void ConvertToHtml(string inputFileName, string outputFileName, Dictionary<string, string> parameters)
        {
            Spire.Presentation.Presentation presentation = new Spire.Presentation.Presentation();

            presentation.LoadFromFile(inputFileName);

            using (var stream = new MemoryStream())
            {
                presentation.SaveToFile(stream, Spire.Presentation.FileFormat.Html);
                string result = System.Text.Encoding.UTF8.GetString(stream.ToArray());
                if(result.IndexOf(replaceValuePptx) >= 0)
                {
                    result = result.Replace(replaceValuePptx, "");
                }
                else
                {
                    result = result.Replace(replaceValue, "");
                }

                File.WriteAllText(outputFileName, result, Encoding.UTF8);
                result = null;
            }
            presentation.Dispose();
        }

        /// <summary>
        /// Convert pdf to html with picture files
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
            Spire.Presentation.Presentation presentation = new Spire.Presentation.Presentation();

            presentation.LoadFromFile(inputFileName);

            using (var stream = new MemoryStream())
            {
                presentation.SaveToFile(stream, Spire.Presentation.FileFormat.Html);
                string result = System.Text.Encoding.UTF8.GetString(stream.ToArray());
                if (result.IndexOf(replaceValuePptx) >= 0)
                {
                    result = result.Replace(replaceValuePptx, "");
                }
                else
                {
                    result = result.Replace(replaceValue, "");
                }
                result = ExtractImages(outputFolder, result);
                File.WriteAllText(outputFolder + "/" + newFileName, result, Encoding.UTF8);
                result = null;
            }
            presentation.Dispose();
        }

        private static string ExtractImages(string outputFolder, string result)
        {
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }
            int index = 1;
            int lastIndex = 0;
            while (true)
            {
                string base64 = getNextB64(result, ref lastIndex);
                if (base64.Equals(""))
                {
                    break;
                }
                using (var memoryStream = new MemoryStream(System.Convert.FromBase64String(base64)))
                {
                    System.Drawing.Image.FromStream(memoryStream).Save(outputFolder + "/img" + index + ".png", System.Drawing.Imaging.ImageFormat.Png);
                    result = result.Replace("xlink:href=\"data:;base64," + base64 + "\"", "xlink:href=\"img" + index + ".png\" ");
                    index++;
                }
            }
            return result;
        }

        private static string getNextB64(string text, ref int lastIndex)
        {
            int imageIndex = text.IndexOf("<image ", lastIndex);
            if (imageIndex < 0)
            {
                return "";
            }
            lastIndex = imageIndex + 7;
            int hrefIndex = text.IndexOf("xlink:href=\"data:;base64,", lastIndex);
            if (hrefIndex < 0)
            {
                return "";
            }
            lastIndex = hrefIndex + 12;
            int endHrefIndex = text.IndexOf("\"", lastIndex);
            if (endHrefIndex < 0)
            {
                return "";
            }
            return text.Substring(hrefIndex + "xlink:href=\"data:;base64,".Length, endHrefIndex - hrefIndex - "xlink:href=\"data:;base64,".Length);
        }
    }
}
