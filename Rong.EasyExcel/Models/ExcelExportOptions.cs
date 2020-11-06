using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel导出配置
    /// </summary>
    public class ExcelExportOptions
    {
        /// <summary>
        /// 工作表名称，默认 Sheet1
        /// </summary>
        [Display(Name = "工作表名称")]
        [StringLength(30, ErrorMessage = "{0}最大长度为{1}")]
        public string SheetName { get; set; }

        /// <summary>
        /// 表头行编号,默认1
        /// <para>从 1 开始</para>
        /// </summary>
        [Display(Name = "表头行编号")]
        [Range(1, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int HeaderRowIndex { get; set; }

        /// <summary>
        /// 数据起始行编号，默认2
        /// <para>从 1 开始</para>
        /// </summary>
        [Display(Name = "数据起始行编号")]
        [Range(2, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int DataRowStartIndex { get; set; }

        /// <summary>
        /// 文件格式，默认 Xlsx
        /// </summary>
        [Display(Name = "文件格式")]
        [EnumDataType(typeof(ExcelTypeEnum), ErrorMessage = "{0}值不存在")]
        public ExcelTypeEnum ExcelType { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelExportOptions()
        {
            SheetName = "Sheet1";
            HeaderRowIndex = 1;
            DataRowStartIndex = 2;
            ExcelType = ExcelTypeEnum.Xlsx;
        }

        /// <summary>
        /// 检查错误
        /// </summary>
        public void CheckError()
        {
            if (DataRowStartIndex <= HeaderRowIndex)
            {
                throw new Exception("【表头行编号】必须小于【数据起始行编号】");
            }
            List<ValidationResult> valid = new List<ValidationResult>();
            var success = Validator.TryValidateObject(this, new ValidationContext(this), valid, true);
            if (!success)
            {
                throw new Exception(valid[0].ErrorMessage);
            }
        }
    }

    /// <summary>
    /// excel文件类型
    /// </summary>
    public enum ExcelTypeEnum
    {
        /// <summary>
        /// Xlsx
        /// <para>Excel为 >= 2007 的版本</para>
        /// <para>application/vnd.openxmlformats-officedocument.spreadsheetml.sheet</para>
        /// </summary>
        Xlsx,
        /// <summary>
        /// Xls
        /// <para>Excel为 &lt;= 2003 的版本</para>
        /// <para>application/vnd.ms-excel</para>
        /// </summary>
        Xls,
    }
}
