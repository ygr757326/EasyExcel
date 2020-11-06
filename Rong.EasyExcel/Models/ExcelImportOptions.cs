using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Rong.EasyExcel.Models
{
    /// <summary>
    /// excel导入配置
    /// </summary>
    public class ExcelImportOptions
    {
        /// <summary>
        /// 工作表编号（默认1）
        /// <para>从 1 开始</para>
        /// <para>0：全部sheet，1：第一个sheet，2：第二个sheet，……</para>
        /// </summary>
        [Display(Name = "工作表编号")]
        [Range(0, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int SheetIndex { get; set; }
        /// <summary>
        /// 表头行编号（默认1）
        /// <para>从 1 开始</para>
        /// </summary>
        [Display(Name = "表头行编号")]
        [Range(1, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int HeaderRowIndex { get; set; }

        /// <summary>
        /// 数据起始行编号（默认2）
        /// <para>从 1 开始</para>
        /// </summary>
        [Display(Name = "数据起始行编号")]
        [Range(2, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int DataRowStartIndex { get; set; }

        /// <summary>
        /// 数据结束行编号（默认最后一行）
        /// <para>从 1 开始</para>
        /// </summary>
        [Display(Name = "数据结束行编号")]
        [Range(2, int.MaxValue, ErrorMessage = "{0}最小值为{1}")]
        public int? DataRowEndIndex { get; set; }

        /// <summary>
        /// 数据校验模式（默认：读取整个工作表后，有无效数据则 抛出异常）
        /// </summary>
        [Display(Name = "数据校验模式")]
        [EnumDataType(typeof(ExcelValidateModeEnum), ErrorMessage = "{0}值不存在")]
        public ExcelValidateModeEnum ValidateMode { get; set; }

        /// <summary>
        /// 构造
        /// </summary>
        public ExcelImportOptions()
        {
            SheetIndex = 1;
            HeaderRowIndex = 1;
            DataRowStartIndex = 2;
            DataRowEndIndex = null;
            ValidateMode = ExcelValidateModeEnum.ThrowSheet;
        }

        /// <summary>
        /// 检查错误
        /// </summary>
        public void CheckError()
        {
            if (DataRowEndIndex < DataRowStartIndex)
            {
                throw new Exception("【数据结束行编号】不能小于【数据起始行编号】");
            }
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
    /// 数据验证不通过的处理方式
    /// </summary>
    public enum ExcelValidateModeEnum
    {
        /// <summary>
        /// 读取某行后，有无效数据则 停止
        /// </summary>
        StopRow,

        /// <summary>
        /// 读取某行后，有无效数据则 抛出异常
        /// </summary>
        ThrowRow,

        /// <summary>
        /// 读取整个工作表后，有无效数据则 停止
        /// </summary>
        StopSheet,

        /// <summary>
        /// 读取整个工作表后，有无效数据则 抛出异常
        /// </summary>
        ThrowSheet,

        /// <summary>
        /// 读取所有工作表后，不抛异常
        /// </summary>
        ReadBook,

        /// <summary>
        /// 读取所有工作表后，有无效数据则 抛出所有异常
        /// </summary>
        ThrowBook
    }
}
