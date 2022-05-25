using System.IO;

namespace Miao.Tools.Excel.Convertor.Utils
{
    /// <summary>
    /// ToCsv
    /// </summary>
    public static class ToCsv
    {
        /// <summary>
        ///  Csv file writes the lines and return the result.
        /// </summary>
        /// <param name="writer"></param>
        /// <param name="data"></param>
        /// <param name="separator"></param>
        public static void CsvWriteLine(TextWriter writer, object[] data, char separator)
        {
            var escapeChars = new[] { separator, '\'', '\n' };
            for (var i = 0; i < data.Length; i++)
            {
                if (i > 0)
                {
                    writer.Write(separator);
                }

                var escape = false;
                var cell = data[i];
                if (cell != null && cell.GetType() == typeof(string))
                {
                    string cellString = cell.ToString() ?? string.Empty;
                    if (cellString.Contains('"'))
                    {
                        escape = true;
                        cellString = cellString.Replace("\"", "\"\"");
                    }
                    else if (cellString.IndexOfAny(escapeChars) >= 0)
                    {
                        escape = true;
                    }
                    cell = cellString;
                }
                if (escape)
                {
                    writer.Write('"');
                }
                writer.Write(cell);
                if (escape)
                {
                    writer.Write('"');
                }
            }
            writer.WriteLine();
        }
    }
}
