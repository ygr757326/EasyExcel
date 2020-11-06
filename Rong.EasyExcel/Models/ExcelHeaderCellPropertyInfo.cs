using System;
using System.Collections.Generic;
using System.Reflection;
using System.Text;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel 表头单元格属性信息
    /// </summary>
    public class ExcelHeaderCellPropertyInfo : ExcelHeaderCellInfo<ExcelHeaderCellProperty>
    {
    }

    /// <summary>
    /// excel 表头单元格属性
    /// </summary>
    public class ExcelHeaderCellProperty : ExcelHeaderCell
    {
        /// <summary>
        /// 列属性信息
        /// </summary>
        public PropertyInfo PropertyInfo { get; set; }
    }
}
