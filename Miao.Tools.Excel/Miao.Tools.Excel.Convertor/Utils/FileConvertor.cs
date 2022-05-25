using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeOpenXml;

namespace Miao.Tools.Excel.Convertor.Utils
{
    /// <summary>
    /// FileConvertor
    /// </summary>
    internal class FileConvertor
    {
        /// <summary>
        /// 转换为文本文件
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="textFile"></param>
        public static void ToTextFile(ExcelWorksheet worksheet, string textFile)
        {
            using var writer = File.CreateText(textFile);
            var startAddress = worksheet.Dimension?.Start;
            var endAddress = worksheet.Dimension?.End;
            if (startAddress == null || endAddress == null)
            {
                return;
            }

            var lines = new List<string>();
            for (int i = startAddress.Row; i <= endAddress.Row; i++)
            {
                var line = new StringBuilder();
                for (int j = startAddress.Column; j <= endAddress.Column; j++)
                {
                    line.Append($"{worksheet.Cells[i, j].Text}\t");
                }
                lines.Add(line.ToString());
            }
            lines.ForEach(x => writer.WriteLine(x));
        }

        /// <summary>
        /// 转换为Csv文件
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="csvFile"></param>
        /// <param name="separator"></param>
        public static void ToCsvFile(ExcelWorksheet worksheet, string csvFile, char separator = ',')
        {
            using var writer = File.CreateText(csvFile);
            var startAddress = worksheet.Dimension?.Start;
            var endAddress = worksheet.Dimension?.End;
            if (startAddress == null || endAddress == null)
            {
                return;
            }

            for (int i = startAddress.Row; i <= endAddress.Row; i++)
            {
                var lineDatas = new List<object>();
                for (int j = startAddress.Column; j <= endAddress.Column; j++)
                {
                    lineDatas.Add(worksheet.Cells[i, j].Text);
                }
                ToCsv.CsvWriteLine(writer, lineDatas.ToArray(), separator);
            }
        }

        /// <summary>
        /// 转换为Html文件
        /// </summary>
        /// <param name="worksheet"></param>
        /// <param name="htmlFile"></param>
        public static void ToHtmlFile(ExcelWorksheet worksheet, string htmlFile)
        {
            using var writer = File.CreateText(htmlFile);
            string htmlString = ToHtml.GetHtml(worksheet);
            writer.Write(htmlString);
        }
    }
}
