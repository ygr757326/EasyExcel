using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel工作表输出
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelSheetDataOutput<T> where T : class, new()
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 工作表编号
        /// <para>从 1 开始</para>
        /// </summary>
        public int SheetIndex { get; set; }

        /// <summary>
        /// 总数据条数
        /// </summary>
        public int TotalCount => Rows?.Count ?? 0;

        /// <summary>
        /// 无效数据数
        /// </summary>
        public int InvalidCount => Rows?.Where(a => !a.IsValid).Count() ?? 0;

        /// <summary>
        /// 有效数据数
        /// </summary>
        public int ValidCount => TotalCount - InvalidCount;

        /// <summary>
        /// 数据集合
        /// </summary>
        public List<ExcelImportRowInfo<T>> Rows { get; set; }
    }

    /// <summary>
    /// 导入行信息
    /// </summary>
    /// <typeparam name="T"></typeparam>
    public class ExcelImportRowInfo<T> where T : class, new()
    {
        /// <summary>
        /// 工作表名称
        /// </summary>
        public string SheetName { get; set; }

        /// <summary>
        /// 工作表编号
        /// <para>从 1 开始</para>
        /// </summary>
        public int SheetIndex { get; set; }

        /// <summary>
        /// 行数据
        /// </summary>
        public T Row { get; set; }

        /// <summary>
        /// 行编号
        /// <para>从 1 开始</para>
        /// </summary>
        public int RowNum { get; set; }

        /// <summary>
        /// 是否有效
        /// </summary>
        public bool IsValid { get; set; }

        /// <summary>
        /// 错误信息（当<see cref="IsValid"/>=false 时才有）
        /// </summary>
        public List<ValidationResult> Errors { get; set; }
    }
}
