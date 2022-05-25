using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Miao.Tools.Excel.Convertor.Utils
{
    /// <summary>
    /// ToHtml
    /// </summary>
    public static class ToHtml
    {
        public static readonly string TableStyle = "border-collapse: collapse;font-family: helvetica, arial, sans-serif;";

        /// <summary>
        /// reader html
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns></returns>
        public static string GetHtml(ExcelWorksheet worksheet)
        {
            var sb = new StringBuilder();
            var cellStyles = new Dictionary<string, string>();

            // Row by row
            var startRow = worksheet.Dimension?.Start?.Row ?? 0;
            var endRow = worksheet.Dimension?.End?.Row ?? 0;
            var startColumn = worksheet.Dimension?.Start?.Column ?? 0;
            var endColumn = worksheet.Dimension?.End?.Column ?? 0;
            for (int row = startRow; row <= endRow; row++)
            {
                if (!worksheet.Row(row).Hidden)
                {
                    sb.AppendLine("<tr>");
                    for (int col = startColumn; col <= endColumn; col++)
                    {
                        if (!worksheet.Column(col).Hidden)
                        {
                            var cell = worksheet.Cells[row, col];
                            int merged = 0;

                            //row is merged
                            if (cell.Merge)
                            {
                                merged = cell.Worksheet.SelectedRange[worksheet.MergedCells[row, col]].Columns;
                            }

                            //11 default font size
                            var x = ProcessCellStyle(cell, cellStyles, worksheet.Column(col).Width, merged);
                            sb.AppendLine(x);
                            if (cell.Merge)
                            {
                                col += (merged - 1);
                            }
                        }
                    }
                    sb.AppendLine("</tr>");
                }
            }

            sb.AppendLine("</table>");
            return string.Format("<table  style=\"{0}>\" data-eth-date=\"{1}\">{2}</table>",
                TableStyle, DateTime.Now, sb.ToString());
        }

        /// <summary>
        /// ProcessCellStyle
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="cellStyles"></param>
        /// <param name="width"></param>
        /// <param name="colSpan"></param>
        /// <returns></returns>
        private static string ProcessCellStyle(ExcelRange cell, Dictionary<string, string> cellStyles, double width = -1, int colSpan = 0)
        {
            cellStyles = new Dictionary<string, string>();
            var sb = new StringBuilder();
            //Border
            PropertyToStyle(cellStyles, "border-top", cell.Style.Border.Top, cellAddress: cell.Address);
            PropertyToStyle(cellStyles, "border-right", cell.Style.Border.Right, cellAddress: cell.Address);
            PropertyToStyle(cellStyles, "border-bottom", cell.Style.Border.Bottom, cellAddress: cell.Address);
            PropertyToStyle(cellStyles, "border-left", cell.Style.Border.Left, cellAddress: cell.Address);

            PropertyToStyle(cellStyles, "text-align", cell.Style.HorizontalAlignment.ToString(), "General");
            PropertyToStyle(cellStyles, "font-weight", cell.Style.Font.Bold == true ? "bold" : "");
            PropertyToStyle(cellStyles, "font-size", cell.Style.Font.Size.ToString(), "11");
            PropertyToStyle(cellStyles, "width", Convert.ToInt16(width * 10));
            PropertyToStyle(cellStyles, "white-space", cell.Style.WrapText == false ? "no-wrap" : "");

            string value = cell.Text;
            if (string.IsNullOrEmpty(value))
            {
                value = "&nbsp;";
            }
            else
            {
                value = System.Net.WebUtility.HtmlEncode(value);
            }
               
            string comment = (cell.Comment != null && cell.Comment.Text != "") ? ("title=\"" + cell.Comment.Text + "\"") : string.Empty;

            if (colSpan > 0)
            {
                sb.AppendFormat("<td style=\"{0}\" eth-cell=\"{1}\" colspan=\"{2}\" {4} >{3}</td>",
                    string.Join(";", cellStyles.Select(x => x.Key + ":" + x.Value)), cell.Address, colSpan, value, comment);
            }
            else
            {
                sb.AppendFormat("<td style=\"{0}\" eth-cell=\"{1}\" {3} >{2}</td>",
                    string.Join(";", cellStyles.Select(x => x.Key + ":" + x.Value)), cell.Address, value, comment);
            }

            return sb.ToString();
        }

        /// <summary>
        /// PropertyToStyle
        /// </summary>
        /// <param name="cellStyles"></param>
        /// <param name="cssproperty"></param>
        /// <param name="value"></param>
        /// <param name="defaultValue"></param>
        /// <param name="cellAddress"></param>
        private static void PropertyToStyle(Dictionary<string, string> cellStyles, string cssproperty, object value, string defaultValue = "", string cellAddress = "")
        {
            if (value == null)
            {
                return;
            }

            string cssItem;
            //borders
            if (value.GetType() == typeof(ExcelBorderItem))
            {
                var temp = (ExcelBorderItem)value;

                if (temp.Style == ExcelBorderStyle.None)
                {
                    return;
                }
                else if (temp.Style == ExcelBorderStyle.Thin)
                {
                    cssItem = "solid 1px ";
                }
                else if (temp.Style == ExcelBorderStyle.Hair || temp.Style == ExcelBorderStyle.Medium)
                {
                    cssItem = "solid 2px ";
                }
                else if (temp.Style == ExcelBorderStyle.Thick)
                {
                    cssItem = "solid 3px ";
                }
                else if (temp.Style == ExcelBorderStyle.Dashed)
                {
                    cssItem = "dashed 1px ";
                }
                else if (temp.Style == ExcelBorderStyle.Dotted)
                {
                    cssItem = "dotted 1px ";
                }
                else
                {
                    cssItem = "solid 2px ";
                }

                //cssItem += GetColor(cellAddress, cssproperty);
                cellStyles.Add(cssproperty, cssItem);
                return;
            }
            else
            {
                cssItem = value.ToString();
            }

            if (cssItem != defaultValue)
            {
                if (cssproperty.Contains("size") || cssproperty.Contains("width"))
                {
                    cellStyles.Add(cssproperty, cssItem.Replace(",", ".") + "px");
                }
                else
                {
                    cellStyles.Add(cssproperty, cssItem);
                }
            }
        }
    }
}
