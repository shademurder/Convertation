using System.IO;
using System.Text;
using System.Collections.Generic;
using System;

namespace CaseAgile.OfficePublisher
{
    class Excel
    {
        private static readonly string replaceValueExcel = "<h2 style=\"color:red\">Evaluation&nbsp;Warning&nbsp;:&nbsp;The&nbsp;document&nbsp;was&nbsp;created&nbsp;with&nbsp;&nbsp;Spire.XLS&nbsp;for&nbsp;.NET</h2>";
        private static readonly string replaceValueWithoutTag = "Evaluation&nbsp;Warning&nbsp;:&nbsp;The&nbsp;document&nbsp;was&nbsp;created&nbsp;with&nbsp;&nbsp;Spire.XLS&nbsp;for&nbsp;.NET";
        private static readonly string replaceValue = "Evaluation Warning : The document was created with  Spire.XLS for .NET";
        /// <summary>
        /// Convert excel to html with picture files
        /// </summary>
        /// <param name="inputFileName">Absolute name of input file</param>
        /// <param name="outputFolder">Folder for saving results</param>
        public static void ConvertToHtml(string inputFileName, string outputFolder, Dictionary<string, string> parameters)
        {
            Spire.Xls.Workbook workbook = new Spire.Xls.Workbook();
            workbook.LoadFromFile(inputFileName);

            int excelMaxRows = GetIntValue(parameters, Parameter.ExcelMaxRows.Value);
            int excelMaxColumns = GetIntValue(parameters, Parameter.ExcelMaxColumns.Value);
            int excelMaxSheetSize = GetIntValue(parameters, Parameter.ExcelMaxSheetSize.Value);

            foreach (Spire.Xls.Worksheet sheet in workbook.Worksheets)
            {
                int lastRow = sheet.LastRow;
                int lastColumn = sheet.LastColumn;

                lastRow = excelMaxRows > 0 && lastRow > excelMaxRows ? excelMaxRows : lastRow;
                lastRow = excelMaxSheetSize > 0 && lastRow > excelMaxSheetSize ? excelMaxSheetSize : lastRow;

                lastColumn = excelMaxColumns > 0 && lastColumn > excelMaxColumns ? excelMaxColumns : lastColumn;
                lastColumn = excelMaxSheetSize > 0 && lastColumn > excelMaxSheetSize ? excelMaxSheetSize : lastColumn;

                sheet.LastRow = lastRow;
                sheet.LastColumn = lastColumn;
                using (var stream = new MemoryStream())
                {
                    sheet.SaveToHtml(stream);
                    string result = System.Text.Encoding.UTF8.GetString(stream.ToArray());

                    if (result.IndexOf(replaceValueExcel) >= 0)
                    {
                        result = result.Replace(replaceValueExcel, "");
                    }
                    else if (result.IndexOf(replaceValueWithoutTag) >= 0)
                    {
                        result = result.Replace(replaceValueWithoutTag, "");
                    }
                    else if (result.IndexOf(replaceValue) >= 0)
                    {
                        result = result.Replace(replaceValue, "");
                    }
                    //result = SetBorder(result, 1);
                    if (workbook.Worksheets.Count > 1)
                    {
                        if (result.IndexOf("<td") >= 0)
                        {
                            File.WriteAllText(outputFolder + "/" + new FileInfo(inputFileName).Name + "." + sheet.Name + ".html", result, Encoding.UTF8);
                        }
                    }
                    else
                    {
                        if (result.IndexOf("<td") >= 0)
                        {
                            File.WriteAllText(outputFolder + "/" + new FileInfo(inputFileName).Name + ".html", result, Encoding.UTF8);
                        }
                    }
                    result = null;
                    sheet.Dispose();
                }
            }
            workbook.Dispose();
        }

        private static string SetBorder(string html, int borderSize)
        {
            return html.Replace("<style type=\"text/css\">", "<style type=\"text/css\">td{border:" + borderSize + "px solid #000000;}");
        }

        private static int GetIntValue(Dictionary<string, string> parameters, string valueName)
        {
            parameters.TryGetValue(valueName, out string value);

            try
            {
                return Int32.Parse(value);
            }
            catch
            {
                return -1;
            }
        }
    }
}
