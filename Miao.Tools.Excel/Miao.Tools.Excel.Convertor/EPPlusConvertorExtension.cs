using System;
using System.Collections.Generic;
using Miao.Tools.Excel.Convertor.Utils;
using OfficeOpenXml;

namespace Miao.Tools.Excel.Convertor
{
    /// <summary>
    /// EPPlus扩展类
    /// </summary>
    public static class EPPlusConvertorExtension
    {
        /// <summary>
        /// 文件转换
        /// </summary>
        /// <param name="worksheet">worksheet</param>
        /// <param name="toFileType">文件转换类型</param>
        /// <param name="filePath">文件路径</param>
        public static void ToFile(this ExcelWorksheet worksheet, ExcelToFileType toFileType, string filePath)
        {
            if (toFileType == ExcelToFileType.Text)
            {
                FileConvertor.ToTextFile(worksheet, filePath);
            }
            else if (toFileType == ExcelToFileType.Csv)
            {
                FileConvertor.ToCsvFile(worksheet, filePath);
            }
            else if(toFileType == ExcelToFileType.Html)
            {
                FileConvertor.ToHtmlFile(worksheet, filePath);
            }
            else
            {
                throw new ArgumentException($"the {toFileType} type is not be supported", nameof(toFileType));
            }
        }
    }
}
