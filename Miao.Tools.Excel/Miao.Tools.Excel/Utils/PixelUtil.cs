using System;
using System.Drawing;
using OfficeOpenXml;

namespace Miao.Tools.Excel.Utils
{
    /// <summary>
    /// PixelUtil
    /// </summary>
    internal static class PixelUtil
    {
        /// <summary>
        /// 获取单元格的宽度(像素)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="config"></param>
        /// <returns></returns>
        public static int GetWidthInPixels(ExcelRange cell, EPPlusConfig config)
        {
            double columnWidth = cell.Worksheet.Column(cell.Start.Column).Width;
            return ToPixelsWidth(cell, columnWidth, config);
        }

        /// <summary>
        /// 获取单元格的高度(像素)
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="config"></param>
        /// <returns></returns>
        public static int GetHeightInPixels(ExcelRange cell, EPPlusConfig config)
        {
            double rowHeight = cell.Worksheet.Row(cell.Start.Row).Height;
            return ToPixelsHeight(rowHeight, config);
        }

        /// <summary>
        /// 转化为像素宽度
        /// </summary>
        /// <param name="cell"></param>
        /// <param name="width"></param>
        /// <param name="config"></param>
        /// <returns></returns>
        public static int ToPixelsWidth(ExcelRange cell, double width, EPPlusConfig config)
        {
            var font = new Font(cell.Style.Font.Name, cell.Style.Font.Size, FontStyle.Regular);
            double pxBaseline = config.WidthPixelsBase;
            if (pxBaseline <= 0)
            {
                pxBaseline = Math.Round(MeasureString("1234567890", font) / 10);
            }
            return (int)(width * pxBaseline);
        }

        /// <summary>
        /// 转化为像素高度
        /// </summary>
        /// <param name="height"></param>
        /// <param name="config"></param>
        /// <returns></returns>
        public static int ToPixelsHeight(double height, EPPlusConfig config)
        {
            using Graphics graphics = Graphics.FromHwnd(IntPtr.Zero);
            float dpiY = graphics.DpiY;
            return (int)(height * (1.0 / config.DefaultDPI) * dpiY);
        }

        /// <summary>
        /// CellPixelWidth
        /// </summary>
        /// <param name="pixels"></param>
        /// <returns></returns>
        public static double CellPixelWidth(double pixels)
        {
            return pixels * 0.75;
        }

        /// <summary>
        /// CellPixelHeight
        /// </summary>
        /// <param name="pixels"></param>
        /// <returns></returns>
        public static double CellPixelHeight(double pixels)
        {
            return pixels * 0.14099 - (pixels * 0.14099 / 100 * -1.30);
        }

        /// <summary>
        /// MeasureString
        /// </summary>
        /// <param name="s"></param>
        /// <param name="font"></param>
        /// <returns></returns>
        public static float MeasureString(string s, Font font)
        {
            using var g = Graphics.FromHwnd(IntPtr.Zero);
            g.TextRenderingHint = System.Drawing.Text.TextRenderingHint.AntiAlias;
            return g.MeasureString(s, font, int.MaxValue, StringFormat.GenericTypographic).Width;
        }
    }
}
