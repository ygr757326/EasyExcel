using System;
using System.Collections.Generic;
using System.Text;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel 导出合并区域信息
    /// </summary>
    public class ExcelExportMergedRegionInfo
    {
        /// <summary>
        /// 起始行（起始下标：0）
        /// </summary>
        public int FromRowIndex { get; set; }

        /// <summary>
        /// 结束行（起始下标：0）
        /// </summary>
        public int ToRowIndex { get; set; }

        /// <summary>
        /// 起始列（起始下标：0）
        /// </summary>
        public int FromColumnIndex { get; set; }

        /// <summary>
        /// 结束列）（起始下标：0）
        /// </summary>
        public int ToColumnIndex { get; set; }

        /// <summary>
        /// 属性名称集合
        /// </summary>
        public string[] PropertyNames { get; set; }

        /// <summary>
        /// 值
        /// </summary>
        public object Value { get; set; }

        #region 方法

        /// <summary>
        /// 值是否相等
        /// </summary>
        /// <returns></returns>
        public bool IsValueEqual(object value)
        {
            if (string.IsNullOrWhiteSpace(value?.ToString()) || string.IsNullOrWhiteSpace(Value?.ToString()))
            {
                return false;
            }
            return Value.Equals(value);
        }

        /// <summary>
        /// 是否能够合并行
        /// </summary>
        /// <returns></returns>
        public bool IsCanMergedRow()
        {
            return FromRowIndex != ToRowIndex && FromColumnIndex == ToColumnIndex;
        }

        /// <summary>
        /// 是否能够合并列
        /// </summary>
        /// <returns></returns>
        public bool IsCanMergedColumn()
        {
            return FromColumnIndex != ToColumnIndex && FromRowIndex == ToRowIndex;
        }

        /// <summary>
        /// 是否在行区域中
        /// </summary>
        /// <returns></returns>
        public bool IsInRangeRow(int rowIndex)
        {
            return rowIndex >= FromRowIndex && rowIndex <= ToRowIndex;
        }

        /// <summary>
        /// 是否在列区域中
        /// </summary>
        /// <returns></returns>
        public bool IsInRangeColumn(int columnIndex)
        {
            return columnIndex >= FromColumnIndex && columnIndex <= ToColumnIndex;
        }

        /// <summary>
        /// 是否在行列区域中
        /// </summary>
        /// <returns></returns>
        public bool IsInRange(int rowIndex, int columnIndex)
        {
            return IsInRangeRow(rowIndex) && IsInRangeColumn(columnIndex);
        }

        /// <summary>
        /// 是否在开始行区域外
        /// </summary>
        /// <returns></returns>
        public bool IsOutRangeRowFrom(int rowIndex)
        {
            return rowIndex < FromRowIndex;
        }

        /// <summary>
        /// 是否在结束行区域外
        /// </summary>
        /// <returns></returns>
        public bool IsOutRangeRowTo(int rowIndex)
        {
            return rowIndex > ToRowIndex;
        }

        /// <summary>
        /// 是否在开始列区域外
        /// </summary>
        /// <returns></returns>
        public bool IsOutRangeColumnFrom(int columnIndex)
        {
            return columnIndex < FromColumnIndex;
        }

        /// <summary>
        /// 是否在结束列区域外
        /// </summary>
        /// <returns></returns>
        public bool IsOutRangeColumnTo(int columnIndex)
        {
            return columnIndex > ToColumnIndex;
        }

        /// <summary>
        /// 与起始结束行是否是相邻行
        /// </summary>
        /// <returns></returns>
        public bool IsSiblingRow(int rowIndex)
        {
            return Math.Abs(rowIndex - ToRowIndex) == 1 || Math.Abs(rowIndex - FromRowIndex) == 1;
        }

        /// <summary>
        /// 起始结束列是否是相邻列
        /// </summary>
        /// <returns></returns>
        public bool IsSiblingColumn(int columnIndex)
        {
            return Math.Abs(columnIndex - ToColumnIndex) == 1 || Math.Abs(columnIndex - FromColumnIndex) == 1;
        }

        /// <summary>
        /// 是否同一行
        /// </summary>
        /// <returns></returns>
        public bool IsSameRow(int rowIndex)
        {
            return rowIndex == FromRowIndex && rowIndex == ToRowIndex;
        }

        /// <summary>
        /// 是否同一列
        /// </summary>
        /// <returns></returns>
        public bool IsSameColumn(int columnIndex)
        {
            return columnIndex == FromColumnIndex && columnIndex == ToColumnIndex;
        }

        #endregion

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelExportMergedRegionInfo()
        {
            PropertyNames = new string[] { };
        }

    }
}
