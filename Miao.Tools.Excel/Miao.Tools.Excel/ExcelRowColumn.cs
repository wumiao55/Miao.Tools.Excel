namespace Miao.Tools.Excel
{
    /// <summary>
    /// ExcelRowColumn
    /// </summary>
    public class ExcelRowColumn
    {
        /// <summary>
        /// 构造方法
        /// </summary>
        public ExcelRowColumn()
        { }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="row"></param>
        /// <param name="column"></param>
        public ExcelRowColumn(int row, int column)
        {
            this.Row = row;
            this.Column = column;
        }

        /// <summary>
        /// 行号
        /// </summary>
        public int Row { get; set; }

        /// <summary>
        /// 列号
        /// </summary>
        public int Column { get; set; }
    }
}
