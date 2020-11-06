using System;
using System.Collections.Generic;
using System.Text;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel 数据行区间下表
    /// </summary>
    public class ExcelDataRowRangeIndex
    {
        /// <summary>
        /// 开始下标（起始为：0）
        /// </summary>
        public int StartIndex { get; set; }
        /// <summary>
        /// 结束下标（起始为：0）
        /// </summary>
        public int EndIndex { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelDataRowRangeIndex()
        {
        }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelDataRowRangeIndex(int startIndex, int endIndex)
        {
            StartIndex = startIndex;
            EndIndex = endIndex;
        }
    }
}
