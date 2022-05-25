using System;

namespace Miao.Tools.Excel
{
    /// <summary>
    /// 配置
    /// </summary>
    public class EPPlusConfig
    {
        /// <summary>
        /// 属性匹配正则
        /// </summary>
        public string PropertyMatchRegex { get; set; } = @"\[(.+)\]";

        /// <summary>
        /// 日期显示格式(如: yyyy-mm-dd HH:mm:ss)
        /// </summary>
        public string DateTimeFormat { get; set; } = "yyyy-mm-dd HH:mm:ss";

        /// <summary>
        /// 默认DPI
        /// </summary>
        public int DefaultDPI { get; set; } = 72;

        /// <summary>
        /// WidthPixelsBase
        /// </summary>
        public double WidthPixelsBase { get; set; }

        /// <summary>
        /// 有图片时的默认尺寸(item1:width; item2:height)
        /// </summary>
        public Tuple<double, double> DefaultSizeWithPicture { get; set; } = new Tuple<double, double>(20.0, 50.0);

        /// <summary>
        /// 是否自动列宽
        /// </summary>
        public bool AutoFitColumn { get; set; }

        /// <summary>
        /// 最大列宽
        /// </summary>
        public double MaxColumnWidth { get; set; } = 100;

        /// <summary>
        /// 设置属性匹配正则
        /// </summary>
        /// <param name="regex">正则表达式</param>
        public void SetPropertyMatchRegex(string regex)
        {
            if (string.IsNullOrEmpty(regex))
            {
                throw new ArgumentNullException(nameof(regex));
            }
            this.PropertyMatchRegex = regex;
        }

        /// <summary>
        /// 设置日期显示格式
        /// </summary>
        /// <param name="format">格式, 如: yyyy-mm-dd HH:mm:ss</param>
        /// <exception cref="ArgumentNullException"></exception>
        public void SetDateTimeFormat(string format)
        {
            if (string.IsNullOrEmpty(format))
            {
                throw new ArgumentNullException(nameof(format));
            }
            this.DateTimeFormat = format;
        }
    }
}
