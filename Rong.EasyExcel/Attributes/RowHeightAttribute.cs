using System;

namespace Rong.EasyExcel.Attributes
{
    /// <summary>
    /// Excel行高特性（导出时用，默认 20）
    /// <para>1.应用在类上</para>
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class RowHeightAttribute : Attribute
    {
        /// <summary>
        /// 表头行高
        /// <para>单位：磅</para>
        /// <para>取值区间：[0-409]</para>
        /// </summary>
        public short HeaderRowHeight { get; set; } = 20;

        /// <summary>
        /// 数据行高
        /// <para>单位：磅</para>
        /// <para>取值区间：[0-409]</para>
        /// </summary>
        public short DataRowHeight { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public RowHeightAttribute()
        {
        }

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="rowHeight">表头/数据 统一行高</param>
        public RowHeightAttribute(short rowHeight)
        {
            HeaderRowHeight = rowHeight;
            DataRowHeight = rowHeight;
        }

        /// <summary>
        /// 构造
        /// </summary>
        public RowHeightAttribute(short headerRowHeight, short dataRowHeight)
        {
            HeaderRowHeight = headerRowHeight;
            DataRowHeight = dataRowHeight;
        }
    }
}
