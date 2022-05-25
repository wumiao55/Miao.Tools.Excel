using System;

namespace Miao.Tools.Excel
{
    /// <summary>
    /// EPPlusCellRange
    /// </summary>
    public class EPPlusCellRange
    {
        private int _fromRow;
        private int _fromColumn;
        private int _toRow;
        private int _toColumn;

        /// <summary>
        /// 构造方法
        /// </summary>
        public EPPlusCellRange() { }

        /// <summary>
        /// 构造方法
        /// </summary>
        /// <param name="fromRow"></param>
        /// <param name="fromColumn"></param>
        /// <param name="toRow"></param>
        /// <param name="toColumn"></param>
        public EPPlusCellRange(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            CheckValue(fromRow);
            CheckValue(fromColumn);
            CheckValue(toRow);
            CheckValue(toColumn);
            _fromRow = fromRow;
            _fromColumn = fromColumn;
            _toRow = toRow;
            _toColumn = toColumn;
        }

        /// <summary>
        /// FromRow
        /// </summary>
        public int FromRow
        {
            get => _fromRow;
            set
            {
                CheckValue(value);
                _fromRow = value;
            }
        }

        /// <summary>
        /// FromColumn
        /// </summary>
        public int FromColumn
        {
            get => _fromColumn;
            set
            {
                CheckValue(value);
                _fromColumn = value;
            }
        }

        /// <summary>
        /// ToRow
        /// </summary>
        public int ToRow
        {
            get => _toRow;
            set
            {
                CheckValue(value);
                _toRow = value;
            }
        }

        /// <summary>
        /// ToColumn
        /// </summary>
        public int ToColumn
        {
            get => _toColumn;
            set
            {
                CheckValue(value);
                _toColumn = value;
            }
        }

        private void CheckValue(int value)
        {
            if (value <= 0)
            {
                throw new InvalidOperationException("EPPlusCellRange's value is illegal, value must be greater than 0");
            }
        }
    }
}
