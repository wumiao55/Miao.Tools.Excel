using System;

namespace Miao.Tools.Excel
{
    /// <summary>
    /// Excel列名称特性
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
    public class ExcelColumnNameAttribute : Attribute
    {
        /// <summary>
        /// 构造方法
        /// </summary>
        public ExcelColumnNameAttribute()
        { }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="columnName">列名称</param>
        public ExcelColumnNameAttribute(string columnName)
        {
            ColumnName = columnName;
        }

        /// <summary>
        /// 列名称
        /// </summary>
        public string ColumnName { get; set; }
    }
}
