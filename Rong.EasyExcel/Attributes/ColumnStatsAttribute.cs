using System;
using System.ComponentModel.DataAnnotations;

namespace Rong.EasyExcel.Attributes
{
    /// <summary>
    /// Excel 列统计特性（导出时用）
    /// <para>1.应用在属性上，可多个</para>
    /// <para>2.单个属性有多个时，须指定向下偏移行数 <see cref="OffsetRow"/></para>
    /// <para>3.支持公式： 求和、平均值、最大值、最小值、计数</para>
    /// </summary>
    [AttributeUsage(AttributeTargets.Property, AllowMultiple = true)]
    public class ColumnStatsAttribute : Attribute
    {
        /// <summary>
        /// 是否展示标签文本（默认: true）
        /// <para>true：当<see cref="Label"/>有内容也不展示</para>
        /// </summary>
        public bool IsShowLabel { get; set; }

        /// <summary>
        /// 标签文本（将在数值的上一行展示）
        /// <para>如果为 null,则默认为：表头+公式名称</para>
        /// </summary>
        public string Label { get; set; }

        /// <summary>
        /// 函数
        /// <para><see cref="FunctionEnum"/></para>
        /// </summary>
        public short Function { get; set; }

        /// <summary>
        /// 将统计展示到某个属性列上
        /// <para>若不指定，则在当前属性列展示</para>
        /// </summary>
        public string ShowOnColumnPropertyName { get; set; }

        /// <summary>
        /// 单位
        /// <para>自动拼接到 <see cref="Label"/> 后的括号中 </para>
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// 向下偏移行数（默认：1）
        /// </summary>
        public int OffsetRow { get; set; } = 1;

        /// <summary>
        /// 构造
        /// </summary>
        public ColumnStatsAttribute()
        {
            IsShowLabel = true;
        }

        /// <summary>
        /// 构造
        /// </summary>
        /// <param name="function">函数 <see cref="FunctionEnum"/></param>
        public ColumnStatsAttribute(short function) : this()
        {
            Function = function;
        }
    }

    /// <summary>
    /// 函数
    /// </summary>
    public enum FunctionEnum : short
    {
        /// <summary>
        /// 空
        /// </summary>
        None,

        /// <summary>
        /// 求和
        /// </summary>
        [Display(Name = "求和")]
        Sum,
        /// <summary>
        /// 平均值
        /// </summary>
        [Display(Name = "平均值")]
        Avg,
        /// <summary>
        /// 计数
        /// </summary>
        [Display(Name = "计数")]
        Count,
        /// <summary>
        /// 最大值
        /// </summary>
        [Display(Name = "最大值")]
        Max,
        /// <summary>
        /// 最小值
        /// </summary>
        [Display(Name = "最小值")]
        Min
    }
}
