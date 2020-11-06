using System;
using System.Collections.Generic;
using System.Text;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel 表头单元格信息
    /// </summary>
    public class ExcelHeaderCellInfo : ExcelHeaderCellInfo<ExcelHeaderCell>
    {
        /// <summary>
        /// 构造
        /// </summary>
        public ExcelHeaderCellInfo(string sheetName, int sheetIndex)
        {
            SheetName = sheetName;
            SheetIndex = sheetIndex;
        }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelHeaderCellInfo(string sheetName, int sheetIndex, List<ExcelHeaderCell> headerCells) : base(sheetName, sheetIndex, headerCells)
        {
        }
    }

    /// <summary>
    /// excel 表头信息
    /// </summary>
    public class ExcelHeaderCellInfo<T> where T : ExcelHeaderCell
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 工作表下标（起始：0）
        /// </summary>
        public int SheetIndex { get; set; }

        /// <summary>
        /// 表头单元格集合
        /// </summary>
        public List<T> HeaderCells { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelHeaderCellInfo()
        {
            HeaderCells = new List<T>();
        }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelHeaderCellInfo(string sheetName, int sheetIndex)
        {
            SheetName = sheetName;
            SheetIndex = sheetIndex;
        }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelHeaderCellInfo(string sheetName, int sheetIndex, List<T> headerCells) : this(sheetName, sheetIndex)
        {
            HeaderCells = headerCells;
        }
    }

    /// <summary>
    /// excel 表头单元格
    /// </summary>
    public class ExcelHeaderCell
    {
        /// <summary>
        /// 表头名称
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 行下标（起始：原值）
        /// </summary>
        public int RowIndex { get; set; }

        /// <summary>
        /// 列下标（起始：原值）
        /// </summary>
        public int ColumnIndex { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelHeaderCell()
        {
        }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelHeaderCell(string name, int rowIndex, int columnIndex)
        {
            Name = name;
            RowIndex = rowIndex;
            ColumnIndex = columnIndex;
        }
    }
}
