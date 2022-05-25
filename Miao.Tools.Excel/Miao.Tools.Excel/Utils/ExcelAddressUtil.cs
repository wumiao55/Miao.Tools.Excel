using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace Miao.Tools.Excel.Utils
{
    /// <summary>
    /// ExcelAddressUtil
    /// </summary>
    internal class ExcelAddressUtil
    {
        /// <summary>
        /// 将Excel地址转换为行列号地址
        /// </summary>
        /// <param name="excelAddress">excel地址,如: A3, D7</param>
        /// <returns></returns>
        public static ExcelRowColumn ConvertExcelRowColumn(string excelAddress)
        {
            string pattern = @"([a-zA-Z]+)([0-9]+)";
            var match = Regex.Match(excelAddress, pattern);
            if (!match.Success || match.Groups.Values.Count() != 3)
            {
                throw new ArgumentException("invalid excel address: " + excelAddress);
            }

            int column = Col(match.Groups.Values.ElementAt(1).Value);
            int row = Convert.ToInt32(match.Groups.Values.ElementAt(2).Value);
            return new ExcelRowColumn(row, column);
        }

        /// <summary>
        /// 检查excel列名合法性
        /// </summary>
        /// <param name="columnName">合法excel列名,如: B, AB, AZ</param>
        /// <returns></returns>
        public static void CheckColumnName(string columnName)
        {
            if (!new Regex("^[a-zA-Z]+$").IsMatch(columnName))
            {
                throw new ArgumentException($"'{columnName}' is invalid column name! valid column name such as: B, AB, AZ", nameof(columnName));
            }
        }

        /// <summary>
        /// 返回Excel列号
        /// </summary>
        /// <param name="columnName">excel列名,如: B, AB, AZ</param>
        /// <returns></returns>
        public static int Col(string columnName)
        {
            var chars = columnName.ToUpper().Reverse().ToList();
            int result = 0;
            for (int i = 0; i < chars.Count; i++)
            {
                var c = chars[i];
                result += ((c - 'A' + 1) * (int)Math.Pow(26, i));
            }
            return result;
        }
    }
}
